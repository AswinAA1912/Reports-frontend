import React, { useMemo, useState, useEffect } from "react";
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
    TextField,
    Button,
    IconButton,
    Tooltip,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Checkbox,
    FormControlLabel,
    FormControl,
    InputLabel,
    Select,
    MenuItem,
    ListItemText,
} from "@mui/material";
import SettingsIcon from "@mui/icons-material/Settings";
import GroupAddIcon from "@mui/icons-material/GroupAdd";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowUpIcon from "@mui/icons-material/KeyboardArrowUp";
import CloseIcon from "@mui/icons-material/Close";
import dayjs from "dayjs";
import PageHeader from "../../Layout/PageHeader";
import AppLayout from "../../Layout/appLayout";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { toast } from "react-toastify";
import { SettingsService } from "../../services/reportSettings.services";
import { DndContext, closestCenter, PointerSensor, useSensor, useSensors } from "@dnd-kit/core";
import { SortableContext, useSortable, verticalListSortingStrategy, arrayMove } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";

// Column config type matching other reports
export interface ColumnConfig {
    key: string;
    label: string;
    enabled: boolean;
    order: number;
    metric?: "qty" | "count";
}

// Initial configurable role columns (all hidden by default as requested)
const DEFAULT_ROLE_COLUMNS: ColumnConfig[] = [
    { key: "ATTEN BY", label: "Atten By", enabled: false, order: 0, metric: "qty" },
    { key: "CREATED", label: "Created", enabled: false, order: 1, metric: "qty" },
    { key: "PRINT", label: "Print", enabled: false, order: 2, metric: "qty" },
    { key: "TAKEN", label: "Taken", enabled: false, order: 3, metric: "qty" },
    { key: "CHECK", label: "Check", enabled: false, order: 4, metric: "qty" },
    { key: "DELIVERY", label: "Delivery", enabled: false, order: 5, metric: "qty" },
    { key: "DRIVER", label: "Driver / Hindi", enabled: false, order: 6, metric: "qty" },
    { key: "SUPERVISOR", label: "Supervisor", enabled: false, order: 7, metric: "qty" },
    { key: "LADIES", label: "Ladies", enabled: false, order: 8, metric: "qty" },
    { key: "Total Tonnage", label: "Total Tonnage", enabled: false, order: 9, metric: "qty" },
    { key: "Count", label: "Count", enabled: false, order: 10, metric: "count" },
];

const getColumnLabel = (col: ColumnConfig) => {
    if (col.key === "Total Tonnage" || col.key === "Count") {
        return col.label;
    }
    const metricLabel = col.metric === "count" ? "Count" : "Qty";
    return `${col.label} (${metricLabel})`;
};

interface SortableColumnRowProps {
    column: ColumnConfig;
    onToggle: (key: string) => void;
    onChangeMetric: (key: string, metric: "qty" | "count") => void;
}

const SortableColumnRow: React.FC<SortableColumnRowProps> = ({ column, onToggle, onChangeMetric }) => {
    const {
        attributes,
        listeners,
        setNodeRef,
        transform,
        transition,
        isDragging
    } = useSortable({ id: column.key });

    const style = {
        transform: CSS.Transform.toString(transform),
        transition,
        opacity: isDragging ? 0.5 : 1,
        backgroundColor: isDragging ? "#f1f5f9" : "transparent",
        borderRadius: "4px",
    };

    return (
        <Box
            ref={setNodeRef}
            style={style}
            display="flex"
            alignItems="center"
            gap={1}
            py={0.5}
            px={1}
        >
            <IconButton
                size="small"
                {...listeners}
                {...attributes}
                sx={{ cursor: "grab" }}
            >
                <DragIndicatorIcon fontSize="small" />
            </IconButton>
            <FormControlLabel
                control={
                    <Checkbox
                        checked={column.enabled}
                        onChange={() => onToggle(column.key)}
                        size="small"
                    />
                }
                label={column.label}
                sx={{ flexGrow: 1, margin: 0 }}
            />
            {column.key !== "Total Tonnage" && column.key !== "Count" && (
                <Select
                    value={column.metric || "qty"}
                    onChange={(e) => onChangeMetric(column.key, e.target.value as "qty" | "count")}
                    size="small"
                    sx={{ minWidth: 80, height: 28, fontSize: "0.75rem" }}
                >
                    <MenuItem value="qty" sx={{ fontSize: "0.75rem" }}>Qty</MenuItem>
                    <MenuItem value="count" sx={{ fontSize: "0.75rem" }}>Count</MenuItem>
                </Select>
            )}
        </Box>
    );
};

interface InvoiceDetail {
    invoiceNo: string;
    date: string;
    customer: string;
    qty: number;
    count: number;
}

interface StaffContribution {
    name: string;
    roleValues: Record<string, { qty: number; count: number }>;
    invoices: InvoiceDetail[];
}

interface VoucherTypeData {
    name: string;
    category: "INWARDS" | "OUTWARDS";
    baseKgs: number;
    staffContributions: StaffContribution[];
}

// Default layout voucher types and dummy data matching user screenshots
const INITIAL_VOUCHERS_DATA: VoucherTypeData[] = [
    // --- INWARDS Vouchers ---
    {
        name: "ON_G.INW - ADJ",
        category: "INWARDS",
        baseKgs: 930,
        staffContributions: [
            {
                name: "BALAMURUGAN.K",
                roleValues: {
                    CREATED: { qty: 930, count: 3 },
                    "Total Tonnage": { qty: 930, count: 3 },
                    Count: { qty: 0, count: 3 },
                },
                invoices: [
                    { invoiceNo: "GIA/000223/26-27", date: "2026-08-08", customer: "Supplier A", qty: 700, count: 1 },
                    { invoiceNo: "GIA/000222/26-27", date: "2026-08-08", customer: "Supplier A", qty: 200, count: 1 },
                    { invoiceNo: "GIA/000224/26-27", date: "2026-08-08", customer: "Supplier A", qty: 30, count: 1 }
                ]
            },
            {
                name: "VELMURUGAN",
                roleValues: {
                    CHECK: { qty: 930, count: 3 },
                    "Total Tonnage": { qty: 930, count: 3 },
                    Count: { qty: 0, count: 3 },
                },
                invoices: [
                    { invoiceNo: "GIA/000223/26-27", date: "2026-08-08", customer: "Supplier A", qty: 700, count: 1 },
                    { invoiceNo: "GIA/000222/26-27", date: "2026-08-08", customer: "Supplier A", qty: 200, count: 1 },
                    { invoiceNo: "GIA/000224/26-27", date: "2026-08-08", customer: "Supplier A", qty: 30, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_M.INW - ADJ",
        category: "INWARDS",
        baseKgs: 37930,
        staffContributions: [
            {
                name: "A.KAVITHA",
                roleValues: {
                    CREATED: { qty: 35280, count: 1 },
                    "Total Tonnage": { qty: 35280, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-003", date: "2026-08-08", customer: "Supplier B", qty: 35280, count: 1 }
                ]
            },
            {
                name: "ARUVAGA KALIDAS",
                roleValues: {
                    CREATED: { qty: 2650, count: 1 },
                    "Total Tonnage": { qty: 2650, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-004", date: "2026-08-08", customer: "Supplier C", qty: 2650, count: 1 }
                ]
            },
            {
                name: "MLM.Guna",
                roleValues: {
                    DRIVER: { qty: 9000, count: 1 },
                    "Total Tonnage": { qty: 9000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-005", date: "2026-08-08", customer: "Supplier D", qty: 9000, count: 1 }
                ]
            },
            {
                name: "PALANICHAMY",
                roleValues: {
                    CHECK: { qty: 3000, count: 1 },
                    "Total Tonnage": { qty: 3000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-006", date: "2026-08-08", customer: "Supplier B", qty: 3000, count: 1 }
                ]
            },
            {
                name: "S.PALANI",
                roleValues: {
                    SUPERVISOR: { qty: 2010, count: 1 },
                    "Total Tonnage": { qty: 2010, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-007", date: "2026-08-08", customer: "Supplier C", qty: 2010, count: 1 }
                ]
            },
            {
                name: "SARAVANAN",
                roleValues: {
                    TAKEN: { qty: 2650, count: 1 },
                    CHECK: { qty: 30270, count: 1 },
                    "Total Tonnage": { qty: 32920, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-008", date: "2026-08-08", customer: "Supplier B", qty: 32920, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_O.INW - ADJ",
        category: "INWARDS",
        baseKgs: 21000,
        staffContributions: [
            {
                name: "BACKIA",
                roleValues: {
                    CREATED: { qty: 21000, count: 1 },
                    "Total Tonnage": { qty: 21000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-009", date: "2026-08-08", customer: "Supplier E", qty: 21000, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_O.M.INW - SJ",
        category: "INWARDS",
        baseKgs: 10235,
        staffContributions: [
            {
                name: "SENTHIL KUMAR N",
                roleValues: {
                    CREATED: { qty: 10235, count: 1 },
                    CHECK: { qty: 10235, count: 1 },
                    "Total Tonnage": { qty: 20470, count: 2 },
                    Count: { qty: 0, count: 2 },
                },
                invoices: [
                    { invoiceNo: "INW-ADJ-010", date: "2026-08-08", customer: "Supplier F", qty: 10235, count: 1 },
                    { invoiceNo: "INW-ADJ-011", date: "2026-08-08", customer: "Supplier F", qty: 10235, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_G.INTRF",
        category: "INWARDS",
        baseKgs: 18576,
        staffContributions: [
            {
                name: "BALAMURUGAN.K",
                roleValues: {
                    CREATED: { qty: 16526, count: 1 },
                    "Total Tonnage": { qty: 16526, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-001", date: "2026-08-08", customer: "Branch A", qty: 16526, count: 1 }
                ]
            },
            {
                name: "CHIDAMBARAM",
                roleValues: {
                    TAKEN: { qty: 5500, count: 1 },
                    "Total Tonnage": { qty: 5500, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-002", date: "2026-08-08", customer: "Branch A", qty: 5500, count: 1 }
                ]
            },
            {
                name: "NANDHA",
                roleValues: {
                    CREATED: { qty: 2050, count: 1 },
                    "Total Tonnage": { qty: 2050, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-003", date: "2026-08-08", customer: "Branch A", qty: 2050, count: 1 }
                ]
            },
            {
                name: "SOUNDARAPANDIAN",
                roleValues: {
                    TAKEN: { qty: 13076, count: 1 },
                    "Total Tonnage": { qty: 13076, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-004", date: "2026-08-08", customer: "Branch B", qty: 13076, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_M.INTRF",
        category: "INWARDS",
        baseKgs: 13520,
        staffContributions: [
            {
                name: "A.KAVITHA",
                roleValues: {
                    CREATED: { qty: 6000, count: 1 },
                    "Total Tonnage": { qty: 6000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-005", date: "2026-08-08", customer: "Branch B", qty: 6000, count: 1 }
                ]
            },
            {
                name: "ARUVAGA KALIDAS",
                roleValues: {
                    CREATED: { qty: 4000, count: 1 },
                    "Total Tonnage": { qty: 4000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-006", date: "2026-08-08", customer: "Branch B", qty: 4000, count: 1 }
                ]
            },
            {
                name: "MADHAN",
                roleValues: {
                    CREATED: { qty: 50, count: 1 },
                    "Total Tonnage": { qty: 50, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-007", date: "2026-08-08", customer: "Branch B", qty: 50, count: 1 }
                ]
            },
            {
                name: "MLM.Guna",
                roleValues: {
                    DRIVER: { qty: 6450, count: 1 },
                    "Total Tonnage": { qty: 6450, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-008", date: "2026-08-08", customer: "Branch C", qty: 6450, count: 1 }
                ]
            },
            {
                name: "MLM.Malaisamy",
                roleValues: {
                    DRIVER: { qty: 3000, count: 1 },
                    "Total Tonnage": { qty: 3000, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-009", date: "2026-08-08", customer: "Branch C", qty: 3000, count: 1 }
                ]
            },
            {
                name: "MLM.SAMY DURAI",
                roleValues: {
                    DRIVER: { qty: 20, count: 1 },
                    "Total Tonnage": { qty: 20, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-010", date: "2026-08-08", customer: "Branch C", qty: 20, count: 1 }
                ]
            },
            {
                name: "N.VIJAYALAKSHMI",
                roleValues: {
                    CREATED: { qty: 3470, count: 1 },
                    "Total Tonnage": { qty: 3470, count: 1 },
                    Count: { qty: 0, count: 1 },
                },
                invoices: [
                    { invoiceNo: "INTRF-011", date: "2026-08-08", customer: "Branch B", qty: 3470, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_G.ADJ - SJ",
        category: "INWARDS",
        baseKgs: 1448,
        staffContributions: [
            {
                name: "SARAVANAN",
                roleValues: {
                    TAKEN: { qty: 654, count: 1 },
                    CHECK: { qty: 793, count: 1 },
                    CREATED: { qty: 1448, count: 1 },
                    "Total Tonnage": { qty: 2895, count: 2 },
                    Count: { qty: 0, count: 2 }
                },
                invoices: [
                    { invoiceNo: "ADJ-001", date: "2026-08-08", customer: "Adj Customer", qty: 1448, count: 2 }
                ]
            }
        ]
    },
    {
        name: "ON_M.ADJ - SJ",
        category: "INWARDS",
        baseKgs: 12720,
        staffContributions: [
            {
                name: "SARAVANAN",
                roleValues: {
                    TAKEN: { qty: 1603, count: 1 },
                    CREATED: { qty: 12720, count: 1 },
                    "Total Tonnage": { qty: 14323, count: 2 },
                    Count: { qty: 0, count: 2 }
                },
                invoices: [
                    { invoiceNo: "ADJ-002", date: "2026-08-08", customer: "Adj Customer", qty: 12720, count: 2 }
                ]
            }
        ]
    },
    {
        name: "ON_M.ATTY",
        category: "INWARDS",
        baseKgs: 32700,
        staffContributions: [
            {
                name: "S.PALANI",
                roleValues: {
                    SUPERVISOR: { qty: 32700, count: 1 },
                    DRIVER: { qty: 16650, count: 1 },
                    "Total Tonnage": { qty: 49350, count: 2 },
                    Count: { qty: 0, count: 2 }
                },
                invoices: [
                    { invoiceNo: "ATTY-001", date: "2026-08-08", customer: "Atty User", qty: 32700, count: 2 }
                ]
            }
        ]
    },
    {
        name: "ON_M.CLEANING",
        category: "INWARDS",
        baseKgs: 8986,
        staffContributions: [
            {
                name: "LATA",
                roleValues: {
                    LADIES: { qty: 5270, count: 1 },
                    SUPERVISOR: { qty: 14256, count: 1 },
                    "Total Tonnage": { qty: 19526, count: 2 },
                    Count: { qty: 0, count: 2 }
                },
                invoices: [
                    { invoiceNo: "CLN-001", date: "2026-08-08", customer: "Cleaning User", qty: 8986, count: 2 }
                ]
            }
        ]
    },
    {
        name: "ON_M.WT-CHK",
        category: "INWARDS",
        baseKgs: 25414,
        staffContributions: [
            {
                name: "SARAVANAN",
                roleValues: {
                    DRIVER: { qty: 8606, count: 1 },
                    SUPERVISOR: { qty: 24904, count: 1 },
                    LADIES: { qty: 23625, count: 1 },
                    "Total Tonnage": { qty: 57135, count: 3 },
                    Count: { qty: 0, count: 3 }
                },
                invoices: [
                    { invoiceNo: "WTC-001", date: "2026-08-08", customer: "Wtc User", qty: 25414, count: 3 }
                ]
            }
        ]
    },
    // --- OUTWARDS Vouchers ---
    {
        name: "1.ON_Sales - M",
        category: "OUTWARDS",
        baseKgs: 60,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 60, count: 1 },
                    "Total Tonnage": { qty: 60, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "SAL-001", date: "2026-08-08", customer: "Retail Customer", qty: 60, count: 1 }
                ]
            }
        ]
    },
    {
        name: "1.ON_M Sales_SM",
        category: "OUTWARDS",
        baseKgs: 23330,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 23330, count: 1 },
                    "Total Tonnage": { qty: 23330, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "SAL-002", date: "2026-08-08", customer: "Retail Customer", qty: 23330, count: 1 }
                ]
            }
        ]
    },
    {
        name: "1.ON_O Sales_SM",
        category: "OUTWARDS",
        baseKgs: 1440,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 1440, count: 1 },
                    "Total Tonnage": { qty: 1440, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "SAL-003", date: "2026-08-08", customer: "Retail Customer", qty: 1440, count: 1 }
                ]
            }
        ]
    },
    {
        name: "1.ON_G Sales_SM",
        category: "OUTWARDS",
        baseKgs: 7066,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 7066, count: 1 },
                    "Total Tonnage": { qty: 7066, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "SAL-004", date: "2026-08-08", customer: "Retail Customer", qty: 7066, count: 1 }
                ]
            }
        ]
    },
    {
        name: "1.ON_G.C.Bill_SM",
        category: "OUTWARDS",
        baseKgs: 6151,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 6151, count: 1 },
                    "Total Tonnage": { qty: 6151, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "SAL-005", date: "2026-08-08", customer: "Retail Customer", qty: 6151, count: 1 }
                ]
            }
        ]
    },
    {
        name: "All Sales",
        category: "OUTWARDS",
        baseKgs: 53299.5,
        staffContributions: [
            {
                name: "SALES_TEAM",
                roleValues: {
                    CREATED: { qty: 53299.5, count: 7 },
                    ATTEN_BY: { qty: 53299.5, count: 7 },
                    PRINT: { qty: 53299.5, count: 7 },
                    TAKEN: { qty: 53299.5, count: 7 },
                    CHECK: { qty: 53299.5, count: 7 },
                    DELIVERY: { qty: 53299.5, count: 7 },
                    DRIVER: { qty: 53299.5, count: 7 },
                    "Total Tonnage": { qty: 373096.5, count: 7 },
                    Count: { qty: 0, count: 7 }
                },
                invoices: [
                    { invoiceNo: "ALL-SAL-001", date: "2026-08-08", customer: "Multiple Customers", qty: 53299.5, count: 7 }
                ]
            }
        ]
    },
    {
        name: "ON_O.M.OUTWARDS",
        category: "OUTWARDS",
        baseKgs: 4000,
        staffContributions: [
            {
                name: "SENTHIL KUMAR N",
                roleValues: {
                    CREATED: { qty: 4000, count: 1 },
                    CHECK: { qty: 4000, count: 1 },
                    "Total Tonnage": { qty: 8000, count: 2 },
                    Count: { qty: 0, count: 2 }
                },
                invoices: [
                    { invoiceNo: "OUT-001", date: "2026-08-08", customer: "Outward Party A", qty: 4000, count: 1 }
                ]
            },
            {
                name: "BALAMURUGAN.K",
                roleValues: {
                    CREATED: { qty: 20, count: 1 },
                    "Total Tonnage": { qty: 20, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-002", date: "2026-08-08", customer: "Outward Party B", qty: 20, count: 1 }
                ]
            },
            {
                name: "VELMURUGAN",
                roleValues: {
                    CHECK: { qty: 20, count: 1 },
                    "Total Tonnage": { qty: 20, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-003", date: "2026-08-08", customer: "Outward Party C", qty: 20, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_G.OUTWARDS",
        category: "OUTWARDS",
        baseKgs: 20,
        staffContributions: [
            {
                name: "AJAY",
                roleValues: {
                    SUPERVISOR: { qty: 2050, count: 1 },
                    "Total Tonnage": { qty: 2050, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-004", date: "2026-08-08", customer: "Outward Party D", qty: 20, count: 1 }
                ]
            },
            {
                name: "KAMARAJ",
                roleValues: {
                    SUPERVISOR: { qty: 3450, count: 1 },
                    "Total Tonnage": { qty: 3450, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-005", date: "2026-08-08", customer: "Outward Party E", qty: 20, count: 1 }
                ]
            },
            {
                name: "KUMAR K",
                roleValues: {
                    SUPERVISOR: { qty: 3616, count: 1 },
                    "Total Tonnage": { qty: 3616, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-006", date: "2026-08-08", customer: "Outward Party F", qty: 20, count: 1 }
                ]
            },
            {
                name: "M.KRISHANAN",
                roleValues: {
                    DRIVER: { qty: 9460, count: 1 },
                    "Total Tonnage": { qty: 9460, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-007", date: "2026-08-08", customer: "Outward Party A", qty: 20, count: 1 }
                ]
            },
            {
                name: "MLM.Malaisamy",
                roleValues: {
                    DRIVER: { qty: 7576, count: 1 },
                    "Total Tonnage": { qty: 7576, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-008", date: "2026-08-08", customer: "Outward Party B", qty: 20, count: 1 }
                ]
            },
            {
                name: "MLM.PRABHU",
                roleValues: {
                    DRIVER: { qty: 2050, count: 1 },
                    "Total Tonnage": { qty: 2050, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-009", date: "2026-08-08", customer: "Outward Party C", qty: 20, count: 1 }
                ]
            },
            {
                name: "MLM.RAJA",
                roleValues: {
                    DRIVER: { qty: 5500, count: 1 },
                    "Total Tonnage": { qty: 5500, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-010", date: "2026-08-08", customer: "Outward Party D", qty: 20, count: 1 }
                ]
            },
            {
                name: "MLM.SAMY DURAI",
                roleValues: {
                    DRIVER: { qty: 3450, count: 1 },
                    "Total Tonnage": { qty: 3450, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-011", date: "2026-08-08", customer: "Outward Party E", qty: 20, count: 1 }
                ]
            },
            {
                name: "SATHYA PRIYA",
                roleValues: {
                    CREATED: { qty: 20076, count: 1 },
                    "Total Tonnage": { qty: 20076, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-012", date: "2026-08-08", customer: "Outward Party F", qty: 20, count: 1 }
                ]
            }
        ]
    },
    {
        name: "ON_M.OUTWARDS",
        category: "OUTWARDS",
        baseKgs: 20076,
        staffContributions: [
            {
                name: "SATHYA PRIYA",
                roleValues: {
                    CREATED: { qty: 20076, count: 1 },
                    "Total Tonnage": { qty: 20076, count: 1 },
                    Count: { qty: 0, count: 1 }
                },
                invoices: [
                    { invoiceNo: "OUT-013", date: "2026-08-08", customer: "Outward Party F", qty: 20076, count: 1 }
                ]
            }
        ]
    }
];

// Initial Group Mapping Config (User groups)
interface GroupConfig {
    name: string;
    parentCategory: "INWARDS" | "OUTWARDS";
    voucherTypes: string[];
}

const INITIAL_GROUPS: GroupConfig[] = [
    {
        name: "PUR / RET",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_G.INW - ADJ", "ON_M.INW - ADJ", "ON_O.INW - ADJ", "ON_O.M.INW - SJ"],
    },
    {
        name: "INT TRF",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_G.INTRF", "ON_M.INTRF"],
    },
    {
        name: "ADJ",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_G.ADJ - SJ", "ON_M.ADJ - SJ"],
    },
    {
        name: "ATTY",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_M.ATTY"],
    },
    {
        name: "CLEANING",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_M.CLEANING"],
    },
    {
        name: "WT.CHECK",
        parentCategory: "INWARDS",
        voucherTypes: ["ON_M.WT-CHK"],
    },
    {
        name: "SALES",
        parentCategory: "OUTWARDS",
        voucherTypes: [
            "1.ON_Sales - M",
            "1.ON_M Sales_SM",
            "1.ON_O Sales_SM",
            "1.ON_G Sales_SM",
            "1.ON_G.C.Bill_SM",
            "All Sales",
        ],
    },
    {
        name: "OUTWARDS / INT TRF",
        parentCategory: "OUTWARDS",
        voucherTypes: ["ON_O.M.OUTWARDS", "ON_G.OUTWARDS", "ON_M.OUTWARDS"],
    },
];

const OverallStaffReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    // UI States
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [drawerOpen, setDrawerOpen] = useState(false);

    // Column Config State
    const [roleColumns, setRoleColumns] = useState<ColumnConfig[]>(() => {
        const saved = sessionStorage.getItem("overallStaffColumns");
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                // Check if the saved columns have the correct format (original keys and a metric field present)
                const isCorrectFormat = parsed.every((c: any) => !c.key.endsWith("_qty") && !c.key.endsWith("_cnt") && ("metric" in c));
                if (isCorrectFormat) {
                    return parsed;
                }
            } catch (e) {
                console.error("Error parsing saved columns", e);
            }
        }
        return DEFAULT_ROLE_COLUMNS;
    });

    const handleChangeMetric = (key: string, metric: "qty" | "count") => {
        setRoleColumns(p => p.map(col => col.key === key ? { ...col, metric } : col));
    };

    useEffect(() => {
        sessionStorage.setItem("overallStaffColumns", JSON.stringify(roleColumns));
    }, [roleColumns]);

    // Drag-and-drop sensors
    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: {
                distance: 5,
            },
        })
    );

    const handleDragEnd = (event: any) => {
        const { active, over } = event;
        if (!over) return;
        if (active.id !== over.id) {
            setRoleColumns((items) => {
                const oldIndex = items.findIndex((item) => item.key === active.id);
                const newIndex = items.findIndex((item) => item.key === over.id);
                const updated = arrayMove(items, oldIndex, newIndex);
                return updated.map((item, idx) => ({ ...item, order: idx }));
            });
        }
    };

    // Grouping Config State
    const [groups, setGroups] = useState<GroupConfig[]>(() => {
        const saved = localStorage.getItem("overallStaffGroups");
        return saved ? JSON.parse(saved) : INITIAL_GROUPS;
    });

    useEffect(() => {
        localStorage.setItem("overallStaffGroups", JSON.stringify(groups));
    }, [groups]);

    // Selected Expanded Voucher Types
    const [expandedVouchers, setExpandedVouchers] = useState<string[]>([]);

    // Selected Expanded Staff Invoices inline
    const [expandedStaff, setExpandedStaff] = useState<string[]>([]);

    // Group Creation Modal Dialog
    const [groupCreateOpen, setGroupCreateOpen] = useState(false);
    const [newGroupName, setNewGroupName] = useState("");
    const [newGroupCategory, setNewGroupCategory] = useState<"INWARDS" | "OUTWARDS">("INWARDS");
    const [newGroupVouchers, setNewGroupVouchers] = useState<string[]>([]);
    const [selectedGroupToEdit, setSelectedGroupToEdit] = useState<string>("");

    // Column Settings Dialog Anchor (Popover/Menu)
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);

    // Template states (Mocking backend templates support)
    const [saveDialogOpen, setSaveDialogOpen] = useState(false);
    const [reportName, setReportName] = useState("");
    const [parentReportName, setParentReportName] = useState("Overall Staff Report");
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [isEditTemplate, setIsEditTemplate] = useState(false);

    // Save Template Handler
    const handleQuickSave = async () => {
        try {
            if (!reportName.trim()) {
                toast.error("Enter Report Name");
                return;
            }

            const payloadColumns = roleColumns.map((c) => ({
                key: c.key,
                label: c.label,
                enabled: c.enabled,
                order: c.order,
                groupBy: 0,
                dataType: "nvarchar"
            }));

            if (selectedTemplateId) {
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 1,
                    reportName: reportName.trim(),
                    columns: payloadColumns
                });
                toast.success("Template Updated Successfully ✅");
            } else {
                const userData = JSON.parse(localStorage.getItem("user") || "{}");
                const createdBy = userData?.id || 0;

                await SettingsService.saveReportSettings({
                    reportName,
                    parentReport: parentReportName,
                    abstractSP: "sp_OverallStaffReport",
                    expandedSP: "sp_OverallStaffReport",
                    abstractColumns: payloadColumns,
                    expandedColumns: payloadColumns,
                    createdBy
                });
                toast.success("Template Saved Successfully ✅");
            }
            setSaveDialogOpen(false);
        } catch (err) {
            console.error(err);
            toast.error("Failed to save template ❌");
        }
    };

    // Load template logic
    const handleLoadTemplate = async (templateId: number) => {
        try {
            setSelectedTemplateId(templateId);
            setIsEditTemplate(true);
            const res = await SettingsService.getReportEditData({ reportId: templateId, typeId: 1 });
            const templateCols = res.data.data.columns || [];

            // Map settings
            const updatedCols = roleColumns.map(col => {
                const matched = templateCols.find((t: any) => t.key === col.key);
                return matched ? { ...col, enabled: matched.enabled } : col;
            });
            setRoleColumns(updatedCols);
            toast.success("Template Loaded Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to load template ❌");
        }
    };

    // Reset Template
    const handleClearTemplate = () => {
        setSelectedTemplateId(null);
        setIsEditTemplate(false);
        setReportName("");
        setRoleColumns(DEFAULT_ROLE_COLUMNS);
    };

    // Toggle Role column visibility
    const handleToggleColumn = (key: string) => {
        setRoleColumns(p => p.map(col => col.key === key ? { ...col, enabled: !col.enabled } : col));
    };

    // Create custom group
    const handleCreateGroup = () => {
        if (!newGroupName.trim()) {
            toast.error("Please enter a group name");
            return;
        }
        if (newGroupVouchers.length === 0) {
            toast.error("Please select at least one voucher type");
            return;
        }

        // Add or replace group
        const newGroup: GroupConfig = {
            name: newGroupName.trim(),
            parentCategory: newGroupCategory,
            voucherTypes: newGroupVouchers
        };

        let updatedGroups = [...groups];

        // If editing an existing group, filter out the old group definition
        if (selectedGroupToEdit && selectedGroupToEdit !== "new") {
            updatedGroups = updatedGroups.filter(g => g.name !== selectedGroupToEdit);
        }

        // Ensure these voucher types are removed from any existing group first
        updatedGroups = updatedGroups.map(g => ({
            ...g,
            voucherTypes: g.voucherTypes.filter(v => !newGroupVouchers.includes(v))
        })).filter(g => g.voucherTypes.length > 0);

        const groupIndex = updatedGroups.findIndex(g => g.name.toLowerCase() === newGroupName.trim().toLowerCase());
        if (groupIndex > -1) {
            updatedGroups[groupIndex] = newGroup;
        } else {
            updatedGroups.push(newGroup);
        }

        setGroups(updatedGroups);
        setNewGroupName("");
        setNewGroupVouchers([]);
        setSelectedGroupToEdit("");
        setGroupCreateOpen(false);
        toast.success(`Group "${newGroup.name}" saved successfully 🎉`);
    };

    const handleCloseGroupDialog = () => {
        setNewGroupName("");
        setNewGroupVouchers([]);
        setSelectedGroupToEdit("");
        setGroupCreateOpen(false);
    };

    // Available voucher types that can be selected in Group Creation
    const allVoucherNames = useMemo(() => INITIAL_VOUCHERS_DATA.map(v => v.name), []);

    // Build hierarchical table data dynamically based on active grouping configuration
    const tableCategories = useMemo(() => {
        // Group the raw vouchers data by the configured group name
        const inwardsGroups: any[] = [];
        const outwardsGroups: any[] = [];

        groups.forEach(group => {
            const matchedVouchers = INITIAL_VOUCHERS_DATA.filter(v => group.voucherTypes.includes(v.name));
            if (matchedVouchers.length === 0) return;

            // Calculate Group Totals
            const groupKgsSum = matchedVouchers.reduce((sum, v) => sum + v.baseKgs, 0);

            const groupData = {
                groupName: group.name,
                groupKgs: groupKgsSum,
                voucherTypes: matchedVouchers.map(v => {
                    // Compute columns sum for this voucher type
                    const totalTonnageSum = v.staffContributions.reduce((sum: number, sc: StaffContribution) => {
                        return sum + (sc.roleValues["Total Tonnage"]?.qty || 0);
                    }, 0);
                    const totalCountSum = v.staffContributions.reduce((sum: number, sc: StaffContribution) => {
                        return sum + (sc.roleValues["Count"]?.count || 0);
                    }, 0);

                    // Compute specific role column sums for header row
                    const roleSums: Record<string, number> = {};
                    roleColumns.forEach(roleCol => {
                        const metric: "qty" | "count" = roleCol.metric || "qty";
                        roleSums[roleCol.key] = v.staffContributions.reduce((sum: number, sc: StaffContribution) => {
                            const scVal = sc.roleValues[roleCol.key]?.[metric] || 0;
                            return sum + scVal;
                        }, 0);
                    });

                    return {
                        name: v.name,
                        baseKgs: v.baseKgs,
                        roleSums,
                        totalTonnage: totalTonnageSum,
                        totalCount: totalCountSum,
                        staff: v.staffContributions
                    };
                })
            };

            if (group.parentCategory === "INWARDS") {
                inwardsGroups.push(groupData);
            } else {
                outwardsGroups.push(groupData);
            }
        });

        // Split inwardsGroups into blocks for display mapping if matching screenshots
        // Block 1 (PUR/RET + INT TRF), Block 2 (ADJ + ATTY + CLEANING + WT.CHECK)
        const inBlock1 = inwardsGroups.filter(g => ["PUR / RET", "INT TRF"].includes(g.groupName));
        const inBlock2 = inwardsGroups.filter(g => ["ADJ", "ATTY", "CLEANING", "WT.CHECK"].includes(g.groupName));

        const categories = [
            {
                name: "INWARDS",
                groups: inBlock1,
                categoryKgs: inBlock1.reduce((sum, g) => sum + g.groupKgs, 0)
            },
            {
                name: "ADJUSTMENTS",
                groups: inBlock2,
                categoryKgs: inBlock2.reduce((sum, g) => sum + g.groupKgs, 0)
            },
            {
                name: "OUTWARDS",
                groups: outwardsGroups,
                categoryKgs: outwardsGroups.reduce((sum, g) => sum + g.groupKgs, 0)
            }
        ].filter(cat => cat.groups.length > 0);

        return categories;
    }, [groups, roleColumns]);

    // Toggles expanded voucher state
    const handleToggleExpandVoucher = (voucherName: string) => {
        setExpandedVouchers(prev =>
            prev.includes(voucherName) ? prev.filter(v => v !== voucherName) : [...prev, voucherName]
        );
    };

    // Calculate Grand Tonnage & Count Totals
    const grandTotals = useMemo(() => {
        let totalKgs = 0;
        let roleTotals: Record<string, number> = {};
        roleColumns.forEach(c => { roleTotals[c.key] = 0; });

        tableCategories.forEach(cat => {
            cat.groups.forEach(g => {
                totalKgs += g.groupKgs;
                g.voucherTypes.forEach((vt: any) => {
                    roleColumns.forEach(c => {
                        roleTotals[c.key] += (vt.roleSums[c.key] || 0);
                    });
                });
            });
        });

        return {
            totalKgs,
            roleTotals
        };
    }, [tableCategories, roleColumns]);

    // Toggle expand staff invoices inline
    const handleToggleExpandStaff = (voucherName: string, staffName: string) => {
        const key = `${voucherName}_${staffName}`;
        setExpandedStaff(prev =>
            prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
        );
    };

    return (
        <>
            <PageHeader
                onExportExcel={() => toast.success("Excel Exported ✅")}
                onExportPDF={() => toast.success("PDF Exported ✅")}
                onReportChange={(template) => {
                    if (!template || !template.Report_Id) {
                        handleClearTemplate();
                    } else {
                        handleLoadTemplate(template.Report_Id);
                    }
                }}
                onQuickSave={(parentName) => {
                    setParentReportName(parentName);
                    setSaveDialogOpen(true);
                }}
                settingsSlot={
                    <Box display="flex" gap={1}>
                        <Tooltip title="Group Creation">
                            <IconButton
                                size="small"
                                onClick={() => setGroupCreateOpen(true)}
                                sx={{
                                    height: 24,
                                    width: 24,
                                    backgroundColor: "#fff",
                                    borderRadius: 0.5,
                                }}
                            >
                                <GroupAddIcon fontSize="small" />
                            </IconButton>
                        </Tooltip>

                        <Tooltip title="Column Settings">
                            <IconButton
                                size="small"
                                onClick={(e) => setSettingsAnchor(e.currentTarget)}
                                sx={{
                                    height: 24,
                                    width: 24,
                                    backgroundColor: "#fff",
                                    borderRadius: 0.5,
                                }}
                            >
                                <SettingsIcon fontSize="small" />
                            </IconButton>
                        </Tooltip>
                    </Box>
                }
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((prev) => !prev)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                onFromDateChange={setFromDate}
                toDate={toDate}
                onToDateChange={setToDate}
                onApply={() => setDrawerOpen(false)}
            />

            <AppLayout fullWidth>
                <Box px={3} pb={4} pt={4}>
                    <Typography variant="subtitle1" fontWeight="bold" color="#1e3a8a" mb={2} sx={{ letterSpacing: 0.5 }}>
                        OVERALL STAFF REPORT {isEditTemplate && ` - ${reportName}`}
                    </Typography>

                    <TableContainer component={Paper} elevation={1} sx={{ borderRadius: 2, border: "1px solid #cbd5e1", overflow: "auto", maxHeight: "calc(100vh - 180px)" }}>
                        <Table size="medium" stickyHeader>
                            <TableHead>
                                <TableRow sx={{ bgcolor: "#1E3A8A" }}>
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }} align="center">
                                        Kgs
                                    </TableCell>
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }} align="center">
                                        Groups
                                    </TableCell>
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }} align="center">
                                        Kgs
                                    </TableCell>
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }}>
                                        Voucher Type
                                    </TableCell>
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }} align="center">
                                        Kgs
                                    </TableCell>

                                    {/* STAFF NAME Header only shown when rows are expanded */}
                                    <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }}>
                                        STAFF NAME
                                    </TableCell>

                                    {/* Dynamically enabled role columns */}
                                    {roleColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(col => (
                                        <TableCell key={col.key} sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.85rem", bgcolor: "#1E3A8A" }} align="center">
                                            {getColumnLabel(col)}
                                        </TableCell>
                                    ))}                                </TableRow>
                            </TableHead>
                            <TableBody>
                                {/* Grand Total Row at the top of the body, styled as sticky */}
                                <TableRow sx={{ borderBottom: "2px solid #1e3a8a" }}>
                                    <TableCell colSpan={3} sx={{ color: "#1e3a8a", fontWeight: 800, fontSize: "0.85rem", position: "sticky", top: 48, zIndex: 100, bgcolor: "#bfdbfe", borderRight: "1px solid #93c5fd" }} align="center">
                                        Total
                                    </TableCell>
                                    <TableCell sx={{ color: "#1e3a8a", fontWeight: 800, fontSize: "0.85rem", position: "sticky", top: 48, zIndex: 100, bgcolor: "#bfdbfe", borderRight: "1px solid #e2e8f0" }}>
                                        Grand Total
                                    </TableCell>
                                    <TableCell sx={{ color: "#1e3a8a", fontWeight: 800, fontSize: "0.85rem", position: "sticky", top: 48, zIndex: 100, bgcolor: "#bfdbfe", borderRight: "1px solid #e2e8f0" }} align="center">
                                        {grandTotals.totalKgs.toLocaleString()}
                                    </TableCell>
                                    <TableCell sx={{ borderRight: "1px solid #93c5fd", position: "sticky", top: 48, zIndex: 100, bgcolor: "#bfdbfe" }} />
                                    {roleColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(col => (
                                        <TableCell key={col.key} sx={{ color: "#1e3a8a", fontWeight: 900, fontSize: "0.85rem", position: "sticky", top: 48, zIndex: 100, bgcolor: "#bfdbfe", borderRight: "1px solid #e2e8f0" }} align="center">
                                            {grandTotals.roleTotals[col.key].toLocaleString()}
                                        </TableCell>
                                    ))}
                                </TableRow>

                                {tableCategories.length > 0 ? (
                                    tableCategories.map((category) => {
                                        // Count visible rows for category rowspan calculation
                                        let categorySpan = 0;
                                        category.groups.forEach((group) => {
                                            group.voucherTypes.forEach((vt: any) => {
                                                categorySpan += 1;
                                                if (expandedVouchers.includes(vt.name)) {
                                                    categorySpan += vt.staff.length;
                                                    vt.staff.forEach((sc: StaffContribution) => {
                                                        const staffKey = `${vt.name}_${sc.name}`;
                                                        if (expandedStaff.includes(staffKey) && sc.invoices) {
                                                            categorySpan += sc.invoices.length;
                                                        }
                                                    });
                                                }
                                            });
                                        });

                                        let isFirstCategoryRow = true;

                                        return category.groups.map((group) => {
                                            // Count visible rows for group rowspan calculation
                                            let groupSpan = 0;
                                            group.voucherTypes.forEach((vt: any) => {
                                                groupSpan += 1;
                                                if (expandedVouchers.includes(vt.name)) {
                                                    groupSpan += vt.staff.length;
                                                    vt.staff.forEach((sc: StaffContribution) => {
                                                        const staffKey = `${vt.name}_${sc.name}`;
                                                        if (expandedStaff.includes(staffKey) && sc.invoices) {
                                                            groupSpan += sc.invoices.length;
                                                        }
                                                    });
                                                }
                                            });

                                            let isFirstGroupRow = true;

                                            return group.voucherTypes.map((vt: any) => {
                                                const isExpanded = expandedVouchers.includes(vt.name);
                                                const hasStaff = vt.staff && vt.staff.length > 0;

                                                const mainRow = (
                                                    <TableRow key={vt.name} sx={{ bgcolor: "#f8fafc", "&:hover": { bgcolor: "#f1f5f9" } }}>
                                                        {/* Yellow Category Spanned Cell */}
                                                        {isFirstCategoryRow && (
                                                            <TableCell
                                                                rowSpan={categorySpan}
                                                                sx={{
                                                                    fontWeight: 800,
                                                                    bgcolor: "#e0f2fe", // Light Blue
                                                                    color: "#0369a1",
                                                                    borderRight: "2px solid #7dd3fc",
                                                                    borderBottom: "2px solid #7dd3fc",
                                                                    verticalAlign: "middle",
                                                                    fontSize: "0.85rem"
                                                                }}
                                                                align="center"
                                                            >
                                                                <Typography variant="body2" fontWeight={800} mb={1}>
                                                                    {category.name}
                                                                </Typography>
                                                                {category.categoryKgs.toLocaleString()}
                                                            </TableCell>
                                                        )}

                                                        {/* Peach Group Name Spanned Cell */}
                                                        {isFirstGroupRow && (
                                                            <TableCell
                                                                rowSpan={groupSpan}
                                                                sx={{
                                                                    fontWeight: 750,
                                                                    bgcolor: "#f0f9ff", // Very Light Blue
                                                                    color: "#0c4a6e",
                                                                    borderRight: "2px solid #bae6fd",
                                                                    borderBottom: "2px solid #bae6fd",
                                                                    verticalAlign: "middle",
                                                                    fontSize: "0.85rem"
                                                                }}
                                                                align="center"
                                                            >
                                                                {group.groupName}
                                                            </TableCell>
                                                        )}

                                                        {/* Peach Group Kgs Spanned Cell */}
                                                        {isFirstGroupRow && (
                                                            <TableCell
                                                                rowSpan={groupSpan}
                                                                sx={{
                                                                    fontWeight: 750,
                                                                    bgcolor: "#f0f9ff", // Very Light Blue
                                                                    color: "#0c4a6e",
                                                                    borderRight: "2px solid #bae6fd",
                                                                    borderBottom: "2px solid #bae6fd",
                                                                    verticalAlign: "middle",
                                                                    fontSize: "0.85rem"
                                                                }}
                                                                align="center"
                                                            >
                                                                {group.groupKgs.toLocaleString()}
                                                            </TableCell>
                                                        )}

                                                        {/* Voucher Type Name */}
                                                        <TableCell
                                                            sx={{
                                                                fontWeight: 700,
                                                                borderRight: "1px solid #cbd5e1",
                                                                color: "#334155",
                                                                cursor: "pointer",
                                                                "&:hover": { color: "#1e3a8a" },
                                                                fontSize: "0.825rem",
                                                                py: 1
                                                            }}
                                                            onClick={() => handleToggleExpandVoucher(vt.name)}
                                                        >
                                                            <Box display="flex" alignItems="center" gap={0.5}>
                                                                {isExpanded ? <KeyboardArrowUpIcon fontSize="small" /> : <KeyboardArrowDownIcon fontSize="small" />}
                                                                {vt.name}
                                                            </Box>
                                                        </TableCell>

                                                        {/* Voucher Type Kgs */}
                                                        <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #cbd5e1", fontSize: "0.825rem", py: 1 }} align="center">
                                                            {vt.baseKgs.toLocaleString()}
                                                        </TableCell>

                                                        {/* Staff Name placeholder for Voucher Type Row */}
                                                        <TableCell sx={{ fontStyle: "italic", color: "text.secondary", borderRight: "1px solid #cbd5e1", py: 1 }} align="center">
                                                            -
                                                        </TableCell>

                                                        {/* Role wise sums for Voucher Type headers */}
                                                        {roleColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(col => {
                                                            const value = vt.roleSums[col.key] || 0;

                                                            return (
                                                                <TableCell key={col.key} sx={{ fontWeight: 800, borderRight: "1px solid #cbd5e1", py: 1, bgcolor: "#f1f5f9" }} align="center">
                                                                    {value > 0 ? value.toLocaleString() : "0"}
                                                                </TableCell>
                                                            );
                                                        })}                                                    </TableRow>
                                                );

                                                isFirstCategoryRow = false;
                                                isFirstGroupRow = false;

                                                // Render Staff nested rows if expanded
                                                const staffRows = isExpanded && hasStaff ? (
                                                    vt.staff.map((sc: StaffContribution) => {
                                                        const isStaffExpanded = expandedStaff.includes(`${vt.name}_${sc.name}`);
                                                        const hasInvoices = sc.invoices && sc.invoices.length > 0;

                                                        const staffRow = (
                                                            <TableRow key={sc.name} sx={{ "&:hover": { bgcolor: "#f8fafc" } }}>
                                                                {/* Empty placeholders for columns spanned by Group & Voucher Type */}
                                                                <TableCell sx={{ borderRight: "1px solid #e2e8f0", py: 0.8 }} />
                                                                <TableCell sx={{ borderRight: "1px solid #e2e8f0", py: 0.8 }} align="center" />

                                                                {/* Clickable Staff Name to Drill down Invoice Details */}
                                                                <TableCell
                                                                    sx={{
                                                                        fontWeight: 650,
                                                                        color: "#1e3a8a",
                                                                        cursor: "pointer",
                                                                        borderRight: "1px solid #e2e8f0",
                                                                        py: 0.8,
                                                                        fontSize: "0.8rem",
                                                                        pl: 4,
                                                                        "&:hover": { textDecoration: "underline" }
                                                                    }}
                                                                    onClick={() => handleToggleExpandStaff(vt.name, sc.name)}
                                                                >
                                                                    <Box display="flex" alignItems="center" gap={0.5}>
                                                                        {isStaffExpanded ? (
                                                                            <KeyboardArrowUpIcon sx={{ fontSize: "0.9rem" }} />
                                                                        ) : (
                                                                            <KeyboardArrowDownIcon sx={{ fontSize: "0.9rem" }} />
                                                                        )}
                                                                        {sc.name}
                                                                    </Box>
                                                                </TableCell>

                                                                {/* Staff role values */}
                                                                {roleColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(col => {
                                                                    const metric: "qty" | "count" = col.metric || "qty";
                                                                    const roleVal = sc.roleValues[col.key]?.[metric] || 0;
                                                                    return (
                                                                        <TableCell key={col.key} sx={{ color: "#475569", borderRight: "1px solid #e2e8f0", py: 0.8, fontSize: "0.8rem" }} align="center">
                                                                            {roleVal > 0 ? roleVal.toLocaleString() : "0"}
                                                                        </TableCell>
                                                                    );
                                                                })}                                                            </TableRow>
                                                        );

                                                        const invoiceRows = isStaffExpanded && hasInvoices ? (
                                                            sc.invoices.map((inv, idx) => (
                                                                <TableRow key={`${sc.name}_${inv.invoiceNo}_${idx}`} sx={{ bgcolor: "#ffffff", "&:hover": { bgcolor: "#f8fafc" } }}>
                                                                    {/* Empty placeholders for category, group, voucher type columns */}
                                                                    <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.6 }} />
                                                                    <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.6 }} align="center" />

                                                                    {/* Invoice number in STAFF NAME column */}
                                                                    <TableCell
                                                                        sx={{
                                                                            color: "#334155",
                                                                            borderRight: "1px solid #cbd5e1",
                                                                            py: 0.6,
                                                                            fontSize: "0.8rem",
                                                                            pl: 6
                                                                        }}
                                                                    >
                                                                        {inv.invoiceNo}
                                                                    </TableCell>

                                                                    {/* Invoice values corresponding to active role columns */}
                                                                    {roleColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(col => {
                                                                        const metric: "qty" | "count" = col.metric || "qty";
                                                                        let val = 0;
                                                                        if (col.key === "Total Tonnage") {
                                                                            val = metric === "qty" ? inv.qty : inv.count;
                                                                        } else if (col.key === "Count") {
                                                                            val = metric === "qty" ? 0 : inv.count;
                                                                        } else {
                                                                            const staffHasRole = sc.roleValues[col.key] && (
                                                                                sc.roleValues[col.key][metric] > 0
                                                                            );
                                                                            if (staffHasRole) {
                                                                                val = metric === "qty" ? inv.qty : inv.count;
                                                                            }
                                                                        }

                                                                        return (
                                                                            <TableCell
                                                                                key={col.key}
                                                                                sx={{
                                                                                    color: "#475569",
                                                                                    borderRight: "1px solid #cbd5e1",
                                                                                    py: 0.6,
                                                                                    fontSize: "0.8rem"
                                                                                }}
                                                                                align="center"
                                                                            >
                                                                                {val.toLocaleString()}
                                                                            </TableCell>
                                                                        );
                                                                    })}                                                                </TableRow>
                                                            ))
                                                        ) : null;

                                                        return (
                                                            <React.Fragment key={sc.name}>
                                                                {staffRow}
                                                                {invoiceRows}
                                                            </React.Fragment>
                                                        );
                                                    })
                                                ) : null;

                                                return (
                                                    <React.Fragment key={vt.name}>
                                                        {mainRow}
                                                        {staffRows}
                                                    </React.Fragment>
                                                );
                                            });
                                        });
                                    })
                                ) : (
                                    <TableRow>
                                        <TableCell colSpan={5 + roleColumns.filter(c => c.enabled).length} align="center" sx={{ py: 6 }}>
                                            No records found. Please create or update groups.
                                        </TableCell>
                                    </TableRow>
                                )}

                                {/* Grand Total moved to top */}
                            </TableBody>
                        </Table>
                    </TableContainer>
                </Box>
            </AppLayout>


            {/* Column Config Settings Dialog popover */}
            <Dialog
                open={Boolean(settingsAnchor)}
                onClose={() => setSettingsAnchor(null)}
                maxWidth="xs"
                fullWidth
            >
                <DialogTitle sx={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    Column Settings
                    <IconButton size="small" onClick={() => setSettingsAnchor(null)}>
                        <CloseIcon />
                    </IconButton>
                </DialogTitle>
                <DialogContent dividers sx={{ maxHeight: "60vh", overflowY: "auto" }}>
                    <Typography variant="subtitle2" fontWeight={700} color="#1E3A8A" mb={1.5}>
                        Drag columns to reorder / Toggle visibility
                    </Typography>

                    <DndContext
                        sensors={sensors}
                        collisionDetection={closestCenter}
                        onDragEnd={handleDragEnd}
                    >
                        <SortableContext
                            items={roleColumns.map(c => c.key)}
                            strategy={verticalListSortingStrategy}
                        >
                            <Box display="flex" flexDirection="column" gap={0.5}>
                                {[...roleColumns]
                                    .sort((a, b) => a.order - b.order)
                                    .map((col) => (
                                        <SortableColumnRow
                                            key={col.key}
                                            column={col}
                                            onToggle={handleToggleColumn}
                                            onChangeMetric={handleChangeMetric}
                                        />
                                    ))}                            </Box>
                        </SortableContext>
                    </DndContext>
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setSettingsAnchor(null)} color="primary" variant="contained">
                        Close
                    </Button>
                </DialogActions>
            </Dialog>            {/* Group Creation Dialog */}
            <Dialog
                open={groupCreateOpen}
                onClose={handleCloseGroupDialog}
                maxWidth="sm"
                fullWidth
            >
                <DialogTitle sx={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    Group Creation
                    <IconButton size="small" onClick={handleCloseGroupDialog}>
                        <CloseIcon />
                    </IconButton>
                </DialogTitle>
                <DialogContent dividers>
                    <FormControl fullWidth size="small" sx={{ mb: 2.5 }}>
                        <InputLabel>Select Group to Edit</InputLabel>
                        <Select
                            value={selectedGroupToEdit}
                            label="Select Group to Edit"
                            onChange={(e) => {
                                const val = e.target.value;
                                setSelectedGroupToEdit(val);
                                if (val && val !== "new") {
                                    const grp = groups.find(g => g.name === val);
                                    if (grp) {
                                        setNewGroupName(grp.name);
                                        setNewGroupCategory(grp.parentCategory);
                                        setNewGroupVouchers(grp.voucherTypes);
                                    }
                                } else {
                                    setNewGroupName("");
                                    setNewGroupCategory("INWARDS");
                                    setNewGroupVouchers([]);
                                }
                            }}
                        >
                            <MenuItem value="new">-- Create New Group --</MenuItem>
                            {groups.map(g => (
                                <MenuItem key={g.name} value={g.name}>{g.name}</MenuItem>
                            ))}
                        </Select>
                    </FormControl>

                    <TextField
                        fullWidth
                        label="Group Name"
                        value={newGroupName}
                        onChange={(e) => setNewGroupName(e.target.value)}
                        placeholder="e.g. PUR / RET, SALES..."
                        sx={{ mb: 2.5 }}
                        size="small"
                    />

                    <FormControl fullWidth size="small" sx={{ mb: 2.5 }}>
                        <InputLabel>Parent Category</InputLabel>
                        <Select
                            value={newGroupCategory}
                            label="Parent Category"
                            onChange={(e) => setNewGroupCategory(e.target.value as any)}
                        >
                            <MenuItem value="INWARDS">INWARDS</MenuItem>
                            <MenuItem value="ADJUSTMENTS">ADJUSTMENTS</MenuItem>
                            <MenuItem value="OUTWARDS">OUTWARDS</MenuItem>
                        </Select>
                    </FormControl>

                    <FormControl fullWidth size="small">
                        <InputLabel>Voucher Types</InputLabel>
                        <Select
                            multiple
                            value={newGroupVouchers}
                            onChange={(e) => setNewGroupVouchers(typeof e.target.value === 'string' ? e.target.value.split(',') : e.target.value)}
                            renderValue={(selected) => selected.join(', ')}
                            label="Voucher Types"
                        >
                            {allVoucherNames.map((name) => (
                                <MenuItem key={name} value={name}>
                                    <Checkbox checked={newGroupVouchers.indexOf(name) > -1} size="small" />
                                    <ListItemText primary={name} />
                                </MenuItem>
                            ))}
                        </Select>
                    </FormControl>
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleCloseGroupDialog}>Cancel</Button>
                    <Button onClick={handleCreateGroup} variant="contained" color="primary">
                        {selectedGroupToEdit && selectedGroupToEdit !== "new" ? "Save Group" : "Create Group"}
                    </Button>
                </DialogActions>
            </Dialog>

            {/* Inline Drill Down replacement - no dialog popup needed */}

            {/* Template Save Dialog */}
            <Dialog
                open={saveDialogOpen}
                onClose={() => setSaveDialogOpen(false)}
            >
                <DialogTitle>
                    {isEditTemplate ? "Edit Template" : "Create Template"}
                </DialogTitle>
                <DialogContent>
                    <TextField
                        fullWidth
                        size="small"
                        label="Report Name"
                        value={reportName}
                        onChange={(e) => setReportName(e.target.value)}
                        sx={{ mt: 1 }}
                    />
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setSaveDialogOpen(false)}>
                        Cancel
                    </Button>
                    <Button variant="contained" onClick={handleQuickSave}>
                        Save
                    </Button>
                </DialogActions>
            </Dialog>
        </>
    );
};

export default OverallStaffReport;
