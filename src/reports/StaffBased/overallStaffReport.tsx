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
    Switch,
    FormControlLabel,
    FormControl,
    InputLabel,
    Select,
    MenuItem,
    Autocomplete,
    CircularProgress,
    Popover,
    RadioGroup,
    Radio,
    FormLabel
} from "@mui/material";
import SettingsIcon from "@mui/icons-material/Settings";
import GroupAddIcon from "@mui/icons-material/GroupAdd";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowUpIcon from "@mui/icons-material/KeyboardArrowUp";
import CloseIcon from "@mui/icons-material/Close";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import PageHeader from "../../Layout/PageHeader";
import AppLayout from "../../Layout/appLayout";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { toast } from "react-toastify";
import { SettingsService } from "../../services/reportSettings.services";
import { employeeReportGroupService, VoucherType } from "../../services/staffBasedReport.services";
import { useAuth } from "../../auth/authContext";
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

// Configurable role columns matching backend stored procedure columns exactly
const DEFAULT_ROLE_COLUMNS: ColumnConfig[] = [
    { key: "Unassigned", label: "Unassigned", enabled: false, order: 0, metric: "qty" },
    { key: "Load Man", label: "Load Man", enabled: false, order: 1, metric: "qty" },
    { key: "Printed", label: "Print", enabled: false, order: 2, metric: "qty" },
    { key: "Supervisor", label: "Supervisor", enabled: false, order: 3, metric: "qty" },
    { key: "Checked", label: "Check / Weight", enabled: false, order: 4, metric: "qty" },
    { key: "Taken", label: "Taken", enabled: false, order: 5, metric: "qty" },
    { key: "Total_Tonnage", label: "Total Tonnage", enabled: false, order: 6, metric: "qty" },
    { key: "Invoice_Count", label: "Count", enabled: false, order: 7, metric: "count" },
];

const discoverRoleColumns = (data: any[], currentCols: ColumnConfig[]): ColumnConfig[] => {
    if (!data || data.length === 0) return currentCols;

    const excludedKeys = new Set([
        "Group_Name",
        "Overall_GroupName",
        "Voucher_Type",
        "Voucher_Ref_Id",
        "Total_Tonnage",
        "Invoice_Count",
    ]);

    const discoveredKeys = new Set<string>();

    data.forEach((row) => {
        if (row.Cost_Category) {
            discoveredKeys.add(row.Cost_Category.trim());
        }
        Object.keys(row).forEach((key) => {
            if (excludedKeys.has(key)) return;
            if (key.endsWith("_Count")) return;
            discoveredKeys.add(key);
        });
    });

    // Merge discovered role columns with existing configuration to preserve order, enabled state, and metric
    const newCols = [...currentCols];

    discoveredKeys.forEach((key) => {
        const exists = newCols.some((c) => c.key.toUpperCase() === key.toUpperCase());
        if (!exists) {
            newCols.push({
                key: key,
                label: key,
                enabled: false,
                order: newCols.length,
                metric: "qty",
            });
        }
    });

    // Make sure Total_Tonnage and Invoice_Count are always present at the end
    const standardKeys = ["Total_Tonnage", "Invoice_Count"];
    const filteredCols = newCols.filter((c) => !standardKeys.includes(c.key));

    // Re-append standard columns to the end if not present
    standardKeys.forEach((key) => {
        let col = newCols.find((c) => c.key === key);
        if (!col) {
            col = {
                key: key,
                label: key === "Total_Tonnage" ? "Total Tonnage" : "Count",
                enabled: false,
                order: filteredCols.length,
                metric: key === "Total_Tonnage" ? "qty" : "count",
            };
        }
        filteredCols.push(col);
    });

    // Normalize order index
    return filteredCols.map((c, index) => ({ ...c, order: index }));
};

const getColumnLabel = (col: ColumnConfig) => {
    if (col.key === "Total_Tonnage" || col.key === "Invoice_Count") {
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
                    <Switch
                        checked={column.enabled}
                        onChange={() => onToggle(column.key)}
                        size="small"
                    />
                }
                label={column.label}
                sx={{ flexGrow: 1, margin: 0 }}
            />
            {column.key !== "Total_Tonnage" && column.key !== "Invoice_Count" && (
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

// Helper to determine the value to render in a staff cell
const getStaffCellValue = (
    sc: any,
    colKey: string,
    metric: "qty" | "count",
    qtyType: "Qty" | "Act_Qty" = "Qty",
    roleColumns?: ColumnConfig[]
) => {
    if (colKey === "Total_Tonnage") {
        return metric === "qty"
            ? (qtyType === "Act_Qty" ? (sc.Total_Act_Qty || 0) : (sc.Total_Qty || 0))
            : (sc.Invoice_Count || 0);
    }
    if (colKey === "Invoice_Count") {
        return metric === "qty" ? 0 : (sc.Invoice_Count || 0);
    }
    if (colKey === "Total_Qty") {
        if (roleColumns) {
            let total = 0;
            roleColumns
                .filter(c => c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count" && c.key !== "Total_Qty" && c.key !== "Total_Count" && c.metric !== "count")
                .forEach(c => {
                    total += getStaffCellValue(sc, c.key, "qty", qtyType);
                });
            return total;
        }
        return 0;
    }
    if (colKey === "Total_Count") {
        if (roleColumns) {
            let total = 0;
            roleColumns
                .filter(c => c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count" && c.key !== "Total_Qty" && c.key !== "Total_Count" && c.metric === "count")
                .forEach(c => {
                    total += getStaffCellValue(sc, c.key, "count", qtyType);
                });
            return total;
        }
        return 0;
    }
    if (sc.roleValues && colKey) {
        const directVal = sc.roleValues[colKey];
        if (directVal !== undefined) {
            return metric === "qty"
                ? (qtyType === "Act_Qty" ? (directVal.actQty || 0) : (directVal.qty || 0))
                : (directVal.count || 0);
        }
        const matchedKey = Object.keys(sc.roleValues).find(k => k.toUpperCase() === colKey.toUpperCase());
        if (matchedKey) {
            const catVal = sc.roleValues[matchedKey];
            return metric === "qty"
                ? (qtyType === "Act_Qty" ? (catVal.actQty || 0) : (catVal.qty || 0))
                : (catVal.count || 0);
        }
    }
    return 0;
};

// Helper to group raw employee data to prevent duplicates
const groupEmployeesList = (rawEmployees: any[]) => {
    const groupedEmployees: any[] = [];
    const employeeMap: Record<string, any> = {};

    rawEmployees.forEach((row: any) => {
        // Group by Cost_Center_Name to prevent duplicate rows for the same person under different IDs (e.g. User ID vs Employee ID)
        const empKey = row.Cost_Center_Name ? String(row.Cost_Center_Name).trim().toUpperCase() : 'unassigned';
        if (!employeeMap[empKey]) {
            employeeMap[empKey] = {
                Emp_Id: row.Emp_Id,
                Cost_Center_Id: row.Cost_Center_Id,
                Cost_Center_Name: row.Cost_Center_Name || 'Unassigned',
                Total_Qty: 0,
                Total_Act_Qty: 0,
                Invoice_Count: 0,
                roleValues: {}
            };
            groupedEmployees.push(employeeMap[empKey]);
        }

        const emp = employeeMap[empKey];

        // Prefer a valid non-zero/non-null employee ID and Cost_Center_Id over a system ID (like "0")
        if (
            (!emp.Emp_Id || emp.Emp_Id === "0" || String(emp.Cost_Center_Id) === "0") &&
            row.Emp_Id &&
            row.Emp_Id !== "0" &&
            String(row.Cost_Center_Id) !== "0"
        ) {
            emp.Emp_Id = row.Emp_Id;
            emp.Cost_Center_Id = row.Cost_Center_Id;
        }

        const category = row.Cost_Category;
        if (category) {
            if (!emp.roleValues[category]) {
                emp.roleValues[category] = { qty: 0, actQty: 0, count: 0 };
            }
            emp.roleValues[category].qty += Number(row.Total_Qty) || 0;
            emp.roleValues[category].actQty += Number(row.Total_Act_Qty ?? row.Act_Qty ?? row.Total_Qty ?? 0) || 0;
            emp.roleValues[category].count += Number(row.Invoice_Count) || 0;
        }
        emp.Total_Qty += Number(row.Total_Qty) || 0;
        emp.Total_Act_Qty += Number(row.Total_Act_Qty ?? row.Act_Qty ?? row.Total_Qty ?? 0) || 0;
        emp.Invoice_Count += Number(row.Invoice_Count) || 0;
    });

    return groupedEmployees;
};

// Helper to determine the value to render in an invoice cell
const getInvoiceCellValue = (
    inv: any,
    colKey: string,
    metric: "qty" | "count",
    qtyType: "Qty" | "Act_Qty" = "Qty",
    roleColumns?: ColumnConfig[]
) => {
    if (colKey === "Total_Tonnage") {
        let total = 0;
        if (inv.roleValues) {
            Object.values(inv.roleValues).forEach((val: any) => {
                total += metric === "qty"
                    ? (qtyType === "Act_Qty" ? (val.actQty || 0) : (val.qty || 0))
                    : (val.count || 0);
            });
        }
        return total;
    }
    if (colKey === "Invoice_Count") {
        let total = 0;
        if (inv.roleValues) {
            Object.values(inv.roleValues).forEach((val: any) => {
                total += metric === "qty" ? 0 : (val.count || 0);
            });
        }
        return total;
    }
    if (colKey === "Total_Qty") {
        if (roleColumns) {
            let total = 0;
            roleColumns
                .filter(c => c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count" && c.key !== "Total_Qty" && c.key !== "Total_Count" && c.metric !== "count")
                .forEach(c => {
                    total += getInvoiceCellValue(inv, c.key, "qty", qtyType);
                });
            return total;
        }
        return 0;
    }
    if (colKey === "Total_Count") {
        if (roleColumns) {
            let total = 0;
            roleColumns
                .filter(c => c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count" && c.key !== "Total_Qty" && c.key !== "Total_Count" && c.metric === "count")
                .forEach(c => {
                    total += getInvoiceCellValue(inv, c.key, "count", qtyType);
                });
            return total;
        }
        return 0;
    }

    // Look up role-specific value from aggregated roleValues
    if (inv.roleValues) {
        const directVal = inv.roleValues[colKey];
        if (directVal !== undefined) {
            return metric === "qty"
                ? (qtyType === "Act_Qty" ? (directVal.actQty || 0) : (directVal.qty || 0))
                : (directVal.count || 0);
        }
        const matchedKey = Object.keys(inv.roleValues).find(k => k.toUpperCase() === colKey.toUpperCase());
        if (matchedKey) {
            const val = inv.roleValues[matchedKey];
            return metric === "qty"
                ? (qtyType === "Act_Qty" ? (val.actQty || 0) : (val.qty || 0))
                : (val.count || 0);
        }
    }
    return 0;
};

const getFilteredStaffByCategories = (staffList: any[], roleColumns: ColumnConfig[], qtyType: "Qty" | "Act_Qty" = "Qty") => {
    const enabledCategoryKeys = roleColumns
        .filter(c => c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count")
        .map(c => c.key.toUpperCase());

    if (enabledCategoryKeys.length === 0) {
        return staffList;
    }

    return staffList.filter((sc: any) => {
        if (!sc.roleValues) return false;
        return Object.keys(sc.roleValues).some(catKey => {
            if (enabledCategoryKeys.includes(catKey.toUpperCase())) {
                const val = sc.roleValues[catKey];
                const qtyVal = qtyType === "Act_Qty" ? (val.actQty || 0) : (val.qty || 0);
                return (qtyVal > 0 || val.count > 0);
            }
            return false;
        });
    });
};

const getFilteredStaffList = (staffList: any[], selectedFilters: string[], roleCols: ColumnConfig[], qtyType: "Qty" | "Act_Qty" = "Qty") => {
    const categoryFiltered = getFilteredStaffByCategories(staffList, roleCols, qtyType);
    return categoryFiltered.filter((sc: any) => {
        if (selectedFilters.length === 0) return true;
        const name = sc.Cost_Center_Name || "Unassigned";
        return selectedFilters.includes(name);
    });
};

// Group Mapping Config (User groups)
interface GroupConfig {
    name: string;
    parentCategory: "INWARDS" | "OUTWARDS" | "ADJUSTMENTS";
    voucherTypes: string[];
}

const OverallStaffReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const { user } = useAuth();
    const companyId = user?.companyId || "default";

    // UI States
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [tempFromDate, setTempFromDate] = useState(today);
    const [tempToDate, setTempToDate] = useState(today);
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [loading, setLoading] = useState(false);

    // Qty vs Act Qty Type State
    const [qtyType, setQtyType] = useState<"Qty" | "Act_Qty">("Qty");

    useEffect(() => {
        if (companyId) {
            const stored = localStorage.getItem(`overall_staff_qty_type_${companyId}`);
            setQtyType(stored === "Act_Qty" ? "Act_Qty" : "Qty");
        }
    }, [companyId]);

    const handleQtyTypeChange = (newType: "Qty" | "Act_Qty") => {
        setQtyType(newType);
        if (companyId) {
            localStorage.setItem(`overall_staff_qty_type_${companyId}`, newType);
        }
    };

    useEffect(() => {
        if (drawerOpen) {
            setTempFromDate(fromDate);
            setTempToDate(toDate);
        }
    }, [drawerOpen, fromDate, toDate]);



    // Live Report Data State
    const [reportData, setReportData] = useState<any[]>([]);

    // Expanded Staff & Invoices Dynamic Data State
    const [staffData, setStaffData] = useState<Record<string, any[]>>({});
    const [invoiceData, setInvoiceData] = useState<Record<string, any[]>>({});

    // Column Config State
    const [roleColumns, setRoleColumns] = useState<ColumnConfig[]>(DEFAULT_ROLE_COLUMNS);
    const [preloading, setPreloading] = useState(false);

    // Template states (Mocking backend templates support)
    const [saveDialogOpen, setSaveDialogOpen] = useState(false);
    const [reportName, setReportName] = useState("");
    const [parentReportName, setParentReportName] = useState("Overall Staff Report");
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [isEditTemplate, setIsEditTemplate] = useState(false);

    const handleChangeMetric = (key: string, metric: "qty" | "count") => {
        setRoleColumns(p => p.map(col => col.key === key ? { ...col, metric } : col));
    };

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
    const [groups, setGroups] = useState<GroupConfig[]>([]);

    // Master Voucher Types loaded from DB
    const [voucherTypes, setVoucherTypes] = useState<VoucherType[]>([]);

    // Involved Staff Filter States
    const [involvedStaff, setInvolvedStaff] = useState<any[]>([]);
    const [selectedStaffFilters, setSelectedStaffFilters] = useState<string[]>([]);
    const [staffFilterAnchor, setStaffFilterAnchor] = useState<null | HTMLElement>(null);
    const [staffSearchQuery, setStaffSearchQuery] = useState("");

    const fetchInvolvedStaff = async () => {
        setPreloading(true);
        try {
            const res = await employeeReportGroupService.getGroupEmployees({
                Fromdate: fromDate,
                Todate: toDate
            });
            if (res.data.success) {
                const rawEmployees = res.data.data || [];
                const uniqueStaff: any[] = [];
                const seen = new Set();
                rawEmployees.forEach((s: any) => {
                    const name = s.Cost_Center_Name || "Unassigned";
                    if (!seen.has(name)) {
                        seen.add(name);
                        uniqueStaff.push(s);
                    }
                });
                setInvolvedStaff(uniqueStaff);
                setRoleColumns(prev => discoverRoleColumns(rawEmployees, prev));
            }
        } catch (err) {
            console.error("Error fetching involved staff list:", err);
        } finally {
            setPreloading(false);
        }
    };

    useEffect(() => {
        if (groups.length > 0) {
            fetchInvolvedStaff();
        }
        setSelectedStaffFilters([]);
    }, [fromDate, toDate, groups]);

    // Pre-load all staff data in background on data/date change to support top-level filter recalculation
    useEffect(() => {
        const loadAllStaffData = async () => {
            if (reportData.length === 0) return;
            setPreloading(true);

            const normalizeCategory = (cat: string) => {
                if (!cat) return "";
                const c = cat.toUpperCase();
                if (c === "PROCESS") return "ADJUSTMENTS";
                return c;
            };

            const newStaffData: Record<string, any[]> = {};
            const accumulatedEmployees: any[] = [];
            let hasNewData = false;

            const promises = reportData.map(async (row) => {
                const parentCategory = normalizeCategory(row.Overall_GroupName);
                const groupName = row.Group_Name || "UNASSIGNED";
                const voucherName = row.Voucher_Type;
                if (!voucherName) return;

                const vKey = `${parentCategory}_${groupName}_${voucherName}`;
                if (!staffData[vKey]) {
                    try {
                        const res = await employeeReportGroupService.getGroupEmployees({
                            Fromdate: fromDate,
                            Todate: toDate,
                            Overall_GroupName: parentCategory === "ADJUSTMENTS" ? "PROCESS" : parentCategory,
                            Group_Name: groupName,
                            Voucher_Type: voucherName
                        });
                        if (res.data.success) {
                            const rawEmployees = res.data.data || [];
                            newStaffData[vKey] = groupEmployeesList(rawEmployees);
                            accumulatedEmployees.push(...rawEmployees);
                            hasNewData = true;
                        }
                    } catch (err) {
                        console.error(`Error pre-loading staff for ${vKey}:`, err);
                    }
                }
            });
            try {
                await Promise.all(promises);
                if (hasNewData) {
                    setStaffData(prev => ({
                        ...prev,
                        ...newStaffData
                    }));
                    if (accumulatedEmployees.length > 0) {
                        setRoleColumns(prev => discoverRoleColumns(accumulatedEmployees, prev));
                    }
                }
            } finally {
                setPreloading(false);
            }
        };

        loadAllStaffData();
    }, [reportData, fromDate, toDate, selectedTemplateId]);

    const handleToggleStaff = (name: string) => {
        setSelectedStaffFilters(prev => {
            if (prev.includes(name)) {
                return prev.filter(n => n !== name);
            } else {
                return [...prev, name];
            }
        });
    };

    const handleToggleSelectAll = () => {
        setSelectedStaffFilters([]);
    };

    // Fetch voucher types and employee report groups from backend on mount
    useEffect(() => {
        const fetchMetadata = async () => {
            try {
                // 1. Fetch Voucher Types
                const vtRes = await employeeReportGroupService.getVoucherTypes();
                if (vtRes.data.success) {
                    setVoucherTypes(vtRes.data.data);
                }

                // 2. Fetch Employee Report Groups
                const ergRes = await employeeReportGroupService.getEmployeeReportGroups();
                if (ergRes.data.success && ergRes.data.data.length > 0) {
                    const groupsMap: Record<string, {
                        name: string;
                        parentCategory: "INWARDS" | "OUTWARDS" | "ADJUSTMENTS";
                        voucherTypes: string[];
                    }> = {};

                    ergRes.data.data.forEach(row => {
                        const name = row.Group_Name;
                        if (!groupsMap[name]) {
                            const catId = Number(row.Overall_GroupId);
                            const parentCategory = catId === 1 ? "INWARDS" : catId === 2 ? "ADJUSTMENTS" : "OUTWARDS";
                            groupsMap[name] = {
                                name,
                                parentCategory,
                                voucherTypes: []
                            };
                        }
                        if (row.Voucher_Type_Name) {
                            groupsMap[name].voucherTypes.push(row.Voucher_Type_Name);
                        }
                    });

                    setGroups(Object.values(groupsMap));
                }
            } catch (err) {
                console.error("Error loading group/voucher metadata:", err);
                toast.error("Failed to load metadata from backend");
            }
        };

        fetchMetadata();
    }, []);

    // Fetch Overall Staff Report Data
    const fetchReportData = async () => {
        setLoading(true);
        try {
            const res = await employeeReportGroupService.getOverallStaffCategorywise({
                Fromdate: fromDate,
                Todate: toDate
            });
            if (res.data.success) {
                const data = res.data.data || [];
                setReportData(data);
                setRoleColumns(prev => discoverRoleColumns(data, prev));
            } else {
                setReportData([]);
            }
        } catch (err) {
            console.error("Error loading report data:", err);
            toast.error("Failed to load report data");
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        if (groups.length > 0) {
            fetchReportData();
        }
    }, [fromDate, toDate, groups]);

    // Selected Expanded Voucher Types (First Expansion level)
    const [expandedVouchers, setExpandedVouchers] = useState<string[]>([]);

    // Selected Expanded Staff Invoices inline (Second Expansion level)
    const [expandedStaff, setExpandedStaff] = useState<string[]>([]);

    useEffect(() => {
        setExpandedVouchers([]);
        setExpandedStaff([]);
        setStaffData({});
        setInvoiceData({});
    }, [fromDate, toDate]);

    // Group Creation Modal Dialog
    const [groupCreateOpen, setGroupCreateOpen] = useState(false);
    const [newGroupName, setNewGroupName] = useState("");
    const [newGroupCategory, setNewGroupCategory] = useState<"INWARDS" | "OUTWARDS" | "ADJUSTMENTS">("INWARDS");
    const [newGroupVouchers, setNewGroupVouchers] = useState<string[]>([]);
    const [selectedGroupToEdit, setSelectedGroupToEdit] = useState<string>("");

    // Column Settings Dialog Anchor (Popover/Menu)
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);



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
                dataType: c.metric || "qty"
            }));

            if (selectedTemplateId) {
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 1,
                    reportName: reportName.trim(),
                    columns: payloadColumns
                });
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 2,
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
                    abstractSP: "SP_Get_EmployeeReport_By_CostCategory_AllModules",
                    expandedSP: "SP_Get_EmployeeReport_By_CostCategory_AllModules",
                    abstractColumns: payloadColumns,
                    expandedColumns: payloadColumns,
                    createdBy
                });
                toast.success("Template Saved Successfully ✅");
            }
            setSaveDialogOpen(false);
            setTimeout(() => {
                window.location.reload();
            }, 300);
        } catch (err) {
            console.error(err);
            toast.error("Failed to save template ❌");
        }
    };

    const safeLocaleString = (val: any) => {
        if (val === undefined || val === null) return "0";
        if (typeof val === "number") return val.toLocaleString();
        return String(val);
    };

    // Load template logic
    const handleLoadTemplate = async (templateId: number) => {
        setPreloading(true);
        try {
            setSelectedTemplateId(templateId);
            setIsEditTemplate(true);
            setExpandedVouchers([]);
            setExpandedStaff([]);
            setStaffData({});
            setInvoiceData({});
            const res = await SettingsService.getReportEditData({ reportId: templateId, typeId: 1 });
            
            const data = res?.data?.data || {};
            if (res.data.success && data.reportInfo) {
                setReportName(data.reportInfo.Report_Name || "");
            }

            const templateCols = Array.isArray(data.abstractColumns)
                ? data.abstractColumns
                : Array.isArray(data.columns)
                    ? data.columns
                    : [];

            // Convert template columns to ColumnConfig format
            const parsedTemplateCols: ColumnConfig[] = templateCols.map((matched: any, idx: number) => {
                const key = matched.key || matched.Key || matched.Column_Name || matched.ColumnName || "";
                const label = matched.label || matched.Label || matched.Alias_Name || key;
                const enabled = matched.enabled ?? matched.Enabled ?? false;
                const order = matched.order ?? matched.Order ?? idx;
                const dataType = matched.dataType || matched.DataType || "qty";
                const metric = (dataType === "qty" || dataType === "count") ? dataType : "qty";

                return {
                    key,
                    label,
                    enabled,
                    order,
                    metric
                };
            });

            // Add any columns from roleColumns that are not in parsedTemplateCols
            const updatedCols = [...parsedTemplateCols];
            roleColumns.forEach(col => {
                const exists = updatedCols.some(c => c.key.toUpperCase() === col.key.toUpperCase());
                if (!exists) {
                    updatedCols.push({
                        ...col,
                        enabled: false,
                        order: updatedCols.length
                    });
                }
            });

            // Re-normalize orders and sort
            updatedCols.sort((a, b) => a.order - b.order);
            const normalizedCols = updatedCols.map((c, index) => ({ ...c, order: index }));

            setRoleColumns(normalizedCols);
            toast.success("Template Loaded Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to load template ❌");
        } finally {
            setPreloading(false);
        }
    };

    // Reset Template
    const handleClearTemplate = () => {
        setSelectedTemplateId(null);
        setIsEditTemplate(false);
        setReportName("");
        setRoleColumns(discoverRoleColumns(involvedStaff, DEFAULT_ROLE_COLUMNS));
        setExpandedVouchers([]);
        setExpandedStaff([]);
        setStaffData({});
        setInvoiceData({});
    };

    // Toggle Role column visibility
    const handleToggleColumn = (key: string) => {
        setRoleColumns(p => p.map(col => col.key === key ? { ...col, enabled: !col.enabled } : col));
    };

    // Create custom group
    const handleCreateGroup = async () => {
        if (!newGroupName.trim()) {
            toast.error("Please enter a group name");
            return;
        }
        if (newGroupVouchers.length === 0) {
            toast.error("Please select at least one voucher type");
            return;
        }

        const categoryIdMap = {
            "INWARDS": 1,
            "ADJUSTMENTS": 2,
            "OUTWARDS": 3
        };
        const overallGroupId = categoryIdMap[newGroupCategory];

        const voucherIds = newGroupVouchers
            .map(vName => voucherTypes.find(vt => vt.label === vName)?.Value)
            .filter((id): id is number => id !== undefined);

        if (voucherIds.length === 0) {
            toast.error("Voucher types could not be mapped to IDs");
            return;
        }

        try {
            const saveRes = await employeeReportGroupService.updateEmployeeReportGroup({
                groupName: newGroupName.trim(),
                overallGroupId,
                voucherIds,
            });

            if (saveRes.data.success) {
                toast.success(`Group "${newGroupName.trim()}" saved successfully 🎉`);

                const ergRes = await employeeReportGroupService.getEmployeeReportGroups();
                if (ergRes.data.success) {
                    const groupsMap: Record<string, {
                        name: string;
                        parentCategory: "INWARDS" | "OUTWARDS" | "ADJUSTMENTS";
                        voucherTypes: string[];
                    }> = {};

                    ergRes.data.data.forEach(row => {
                        const name = row.Group_Name;
                        if (!groupsMap[name]) {
                            const catId = Number(row.Overall_GroupId);
                            const parentCategory = catId === 1 ? "INWARDS" : catId === 2 ? "ADJUSTMENTS" : "OUTWARDS";
                            groupsMap[name] = {
                                name,
                                parentCategory,
                                voucherTypes: []
                            };
                        }
                        if (row.Voucher_Type_Name) {
                            groupsMap[name].voucherTypes.push(row.Voucher_Type_Name);
                        }
                    });

                    setGroups(Object.values(groupsMap));
                }

                setNewGroupName("");
                setNewGroupVouchers([]);
                setSelectedGroupToEdit("");
                setGroupCreateOpen(false);
            } else {
                toast.error("Failed to save group to server ❌");
            }
        } catch (err) {
            console.error("Error saving group:", err);
            toast.error("Failed to save group to server ❌");
        }
    };

    const handleCloseGroupDialog = () => {
        setNewGroupName("");
        setNewGroupVouchers([]);
        setSelectedGroupToEdit("");
        setGroupCreateOpen(false);
    };

    // Available voucher types that can be selected in Group Creation
    const allVoucherNames = useMemo(() => voucherTypes.map(v => v.label), [voucherTypes]);

    // Build hierarchical table data dynamically based on active grouping config and live reportData
    const tableCategories = useMemo(() => {
        const normalizeCategory = (cat: string) => {
            if (!cat) return "";
            const c = cat.toUpperCase();
            if (c === "PROCESS") return "ADJUSTMENTS";
            return c;
        };

        const catList = ["INWARDS", "ADJUSTMENTS", "OUTWARDS"];
        const categoryGroupsMap: Record<string, any[]> = {
            INWARDS: [],
            ADJUSTMENTS: [],
            OUTWARDS: []
        };

        // 1. Process defined groups
        groups.forEach(group => {
            const matchedRows = reportData.filter(row =>
                row.Group_Name === group.name &&
                normalizeCategory(row.Overall_GroupName) === group.parentCategory &&
                row.Voucher_Type
            );
            if (matchedRows.length === 0) return;

            const voucherTypesData = matchedRows.map(r => {
                const vKey = `${group.parentCategory}_${group.name}_${r.Voucher_Type}`;
                const staffList = staffData[vKey] || [];
                const filteredStaffList = getFilteredStaffList(staffList, selectedStaffFilters, roleColumns, qtyType);

                const roleSums: Record<string, number> = {};
                let totalTonnage = 0;
                let totalCount = 0;

                roleColumns.filter(roleCol => roleCol.enabled && roleCol.key !== "Total_Tonnage" && roleCol.key !== "Invoice_Count" && roleCol.key !== "Total_Qty" && roleCol.key !== "Total_Count").forEach(roleCol => {
                    let sumVal = 0;
                    filteredStaffList.forEach(sc => {
                        sumVal += getStaffCellValue(sc, roleCol.key, roleCol.metric || "qty", qtyType, roleColumns);
                    });
                    roleSums[roleCol.key] = sumVal;
                });

                let enabledQtySum = 0;
                let enabledCountSum = 0;
                roleColumns.forEach(col => {
                    if (col.enabled && col.key !== "Total_Tonnage" && col.key !== "Invoice_Count" && col.key !== "Total_Qty" && col.key !== "Total_Count") {
                        if (col.metric === "count") {
                            enabledCountSum += roleSums[col.key] || 0;
                        } else {
                            enabledQtySum += roleSums[col.key] || 0;
                        }
                    }
                });

                roleSums["Total_Qty"] = enabledQtySum;
                roleSums["Total_Count"] = enabledCountSum;

                roleSums["Total_Tonnage"] = qtyType === "Act_Qty"
                    ? (r.Total_Act_Tonnage ?? r.Act_Total_Tonnage ?? r.Total_Act_Qty ?? r.Total_Tonnage ?? 0)
                    : (r.Total_Tonnage || 0);
                roleSums["Invoice_Count"] = r.Invoice_Count || 0;

                totalTonnage = roleSums["Total_Tonnage"];
                totalCount = roleSums["Invoice_Count"];

                return {
                    name: r.Voucher_Type,
                    baseKgs: totalTonnage,
                    roleSums,
                    totalTonnage,
                    totalCount,
                    staff: staffList
                };
            });

            const groupKgsSum = voucherTypesData.reduce((sum, vt) => sum + vt.totalTonnage, 0);

            const groupData = {
                groupName: group.name,
                groupKgs: groupKgsSum,
                voucherTypes: voucherTypesData
            };

            if (categoryGroupsMap[group.parentCategory]) {
                categoryGroupsMap[group.parentCategory].push(groupData);
            }
        });

        // 2. Process Unassigned / Null groups for each category
        catList.forEach(cat => {
            const definedGroupNames = groups.filter(g => g.parentCategory === cat).map(g => g.name);

            // Find rows for this category that don't belong to any defined group
            const unassignedRows = reportData.filter(row =>
                normalizeCategory(row.Overall_GroupName) === cat &&
                row.Voucher_Type &&
                (row.Group_Name === null || row.Group_Name === undefined || !definedGroupNames.includes(row.Group_Name))
            );

            if (unassignedRows.length > 0) {
                const voucherTypesData = unassignedRows.map(r => {
                    const vKey = `${cat}_UNASSIGNED_${r.Voucher_Type}`;
                    const staffList = staffData[vKey] || [];
                    const filteredStaffList = getFilteredStaffList(staffList, selectedStaffFilters, roleColumns, qtyType);

                    const roleSums: Record<string, number> = {};
                    let totalTonnage = 0;
                    let totalCount = 0;

                    roleColumns.filter(roleCol => roleCol.enabled && roleCol.key !== "Total_Tonnage" && roleCol.key !== "Invoice_Count" && roleCol.key !== "Total_Qty" && roleCol.key !== "Total_Count").forEach(roleCol => {
                        let sumVal = 0;
                        filteredStaffList.forEach(sc => {
                            sumVal += getStaffCellValue(sc, roleCol.key, roleCol.metric || "qty", qtyType, roleColumns);
                        });
                        roleSums[roleCol.key] = sumVal;
                    });

                    let enabledQtySum = 0;
                    let enabledCountSum = 0;
                    roleColumns.forEach(col => {
                        if (col.enabled && col.key !== "Total_Tonnage" && col.key !== "Invoice_Count" && col.key !== "Total_Qty" && col.key !== "Total_Count") {
                            if (col.metric === "count") {
                                enabledCountSum += roleSums[col.key] || 0;
                            } else {
                                enabledQtySum += roleSums[col.key] || 0;
                            }
                        }
                    });

                    roleSums["Total_Qty"] = enabledQtySum;
                    roleSums["Total_Count"] = enabledCountSum;

                    roleSums["Total_Tonnage"] = qtyType === "Act_Qty"
                        ? (r.Total_Act_Tonnage ?? r.Act_Total_Tonnage ?? r.Total_Act_Qty ?? r.Total_Tonnage ?? 0)
                        : (r.Total_Tonnage || 0);
                    roleSums["Invoice_Count"] = r.Invoice_Count || 0;

                    totalTonnage = roleSums["Total_Tonnage"];
                    totalCount = roleSums["Invoice_Count"];

                    return {
                        name: r.Voucher_Type,
                        baseKgs: totalTonnage,
                        roleSums,
                        totalTonnage,
                        totalCount,
                        staff: staffList
                    };
                });

                const groupKgsSum = voucherTypesData.reduce((sum, vt) => sum + vt.totalTonnage, 0);
                const unassignedGroupData = {
                    groupName: "UNASSIGNED",
                    groupKgs: groupKgsSum,
                    voucherTypes: voucherTypesData
                };
                categoryGroupsMap[cat].push(unassignedGroupData);
            }
        });

        // 3. Construct categories output
        const categories = catList.map(cat => ({
            name: cat,
            groups: categoryGroupsMap[cat],
            categoryKgs: categoryGroupsMap[cat].reduce((sum, g) => sum + g.groupKgs, 0)
        })).filter(cat => cat.groups.length > 0);

        return categories;
    }, [groups, reportData, roleColumns, staffData, selectedStaffFilters, qtyType]);

    const gridStaffNames = useMemo(() => {
        const names = new Set<string>();
        tableCategories.forEach(cat => {
            cat.groups.forEach(g => {
                g.voucherTypes.forEach((vt: any) => {
                    const filteredStaff = getFilteredStaffList(vt.staff || [], [], roleColumns, qtyType);
                    filteredStaff.forEach((sc: any) => {
                        if (sc.Cost_Center_Name) {
                            names.add(sc.Cost_Center_Name);
                        }
                    });
                });
            });
        });
        return Array.from(names).sort();
    }, [tableCategories, roleColumns, qtyType]);

    const filteredGridStaffNames = useMemo(() => {
        return gridStaffNames.filter(name =>
            name.toLowerCase().includes(staffSearchQuery.toLowerCase())
        );
    }, [gridStaffNames, staffSearchQuery]);

    // Toggles expanded voucher state & fetches group employees dynamically
    const handleToggleExpandVoucher = async (parentCategory: string, groupName: string, voucherName: string) => {
        const key = `${parentCategory}_${groupName}_${voucherName}`;
        const isExpanded = expandedVouchers.includes(key);

        if (!isExpanded) {
            setExpandedVouchers(prev => [...prev, key]);
            if (!staffData[key]) {
                try {
                    const res = await employeeReportGroupService.getGroupEmployees({
                        Fromdate: fromDate,
                        Todate: toDate,
                        Overall_GroupName: parentCategory === "ADJUSTMENTS" ? "PROCESS" : parentCategory,
                        Group_Name: groupName,
                        Voucher_Type: voucherName
                    });
                    if (res.data.success) {
                        setStaffData(prev => ({
                            ...prev,
                            [key]: groupEmployeesList(res.data.data || [])
                        }));
                    }
                } catch (err) {
                    console.error("Error loading group employees:", err);
                    toast.error("Failed to load staff list");
                }
            }
        } else {
            setExpandedVouchers(prev => prev.filter(v => v !== key));
        }
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

        // Calculate grand totals for Total_Qty and Total_Count based on enabled columns
        let enabledQtyGrandSum = 0;
        let enabledCountGrandSum = 0;
        roleColumns.forEach(c => {
            if (c.enabled && c.key !== "Total_Tonnage" && c.key !== "Invoice_Count" && c.key !== "Total_Qty" && c.key !== "Total_Count") {
                if (c.metric === "count") {
                    enabledCountGrandSum += roleTotals[c.key] || 0;
                } else {
                    enabledQtyGrandSum += roleTotals[c.key] || 0;
                }
            }
        });

        roleTotals["Total_Qty"] = enabledQtyGrandSum;
        roleTotals["Total_Count"] = enabledCountGrandSum;

        let dynamicQtyTotal = 0;
        let dynamicCountTotal = 0;
        roleColumns.filter(c => c.enabled).forEach(c => {
            const val = roleTotals[c.key] || 0;
            if (c.metric === "count") {
                dynamicCountTotal += val;
            } else {
                dynamicQtyTotal += val;
            }
        });

        return {
            totalKgs,
            roleTotals,
            dynamicQtyTotal,
            dynamicCountTotal
        };
    }, [tableCategories, roleColumns]);

    // Toggle expand staff invoices inline & fetches invoices dynamically
    const handleToggleExpandStaff = async (parentCategory: string, groupName: string, voucherName: string, empId: number, _staffName: string) => {
        const vKey = `${parentCategory}_${groupName}_${voucherName}`;
        const staffKey = `${vKey}_${empId || 'unassigned'}`;
        const isExpanded = expandedStaff.includes(staffKey);

        if (!isExpanded) {
            setExpandedStaff(prev => [...prev, staffKey]);
            if (!invoiceData[staffKey]) {
                try {
                    const res = await employeeReportGroupService.getEmployeeInvoices({
                        Fromdate: fromDate,
                        Todate: toDate,
                        Overall_GroupName: parentCategory === "ADJUSTMENTS" ? "PROCESS" : parentCategory,
                        Group_Name: groupName,
                        Voucher_Type: voucherName,
                        Emp_Id: empId || undefined,
                        Emp_Id_Is_Unassigned: empId ? 0 : 1
                    });
                    if (res.data.success) {
                        // Group raw invoices by Inv_No (or fallback to Voucher_Ref_Id if Inv_No is missing) to prevent duplicate invoice groupings
                        const rawInvoices = res.data.data || [];
                        const groupedInvoices: any[] = [];
                        const groupMap: Record<string, any> = {};

                        rawInvoices.forEach((row: any) => {
                            const refId = row.Inv_No ? String(row.Inv_No).trim().toUpperCase() : String(row.Voucher_Ref_Id);
                            if (!groupMap[refId]) {
                                groupMap[refId] = {
                                    Voucher_Ref_Id: row.Voucher_Ref_Id,
                                    Inv_No: row.Inv_No,
                                    Voucher_Type: row.Voucher_Type,
                                    Group_Name: row.Group_Name,
                                    Overall_GroupName: row.Overall_GroupName,
                                    Voucher_Date: row.Voucher_Date,
                                    Cost_Center_Id: row.Cost_Center_Id,
                                    Cost_Center_Name: row.Cost_Center_Name,
                                    roleValues: {}
                                };
                                groupedInvoices.push(groupMap[refId]);
                            }

                            const inv = groupMap[refId];
                            const category = row.Cost_Category;
                            if (category) {
                                if (!inv.roleValues[category]) {
                                    inv.roleValues[category] = { qty: 0, actQty: 0, count: 0 };
                                }
                                inv.roleValues[category].qty += Number(row.Bill_Qty) || 0;
                                inv.roleValues[category].actQty += Number(row.Bill_Act_Qty ?? row.Act_Qty ?? row.Bill_Qty ?? 0) || 0;
                                inv.roleValues[category].count = 1;
                            }
                        });

                        setInvoiceData(prev => ({
                            ...prev,
                            [staffKey]: groupedInvoices
                        }));
                    }
                } catch (err) {
                    console.error("Error loading employee invoices:", err);
                    toast.error("Failed to load invoices list");
                }
            }
        } else {
            setExpandedStaff(prev => prev.filter(k => k !== staffKey));
        }
    };

    const headerCellHeight = 38;
    const headerCellSx = {
        color: "#ffffff",
        fontWeight: 700,
        height: headerCellHeight,
        py: 0,
        borderRight: "1px solid #2448b2",
        fontSize: "0.75rem",
        bgcolor: "#1E3A8A",
        whiteSpace: "nowrap",
        position: "sticky",
        top: 0,
        zIndex: 3
    };

    const totalCellSx = {
        color: "#1e3a8a",
        fontWeight: 800,
        fontSize: "0.725rem",
        py: 1,
        bgcolor: "#bfdbfe",
        borderRight: "1px solid #93c5fd",
        whiteSpace: "nowrap",
        position: "sticky",
        top: headerCellHeight,
        zIndex: 2
    };

    const handleExportExcel = () => {
        const enabledRoles = roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order);
        const rows: any[] = [];

        tableCategories.forEach((category) => {
            category.groups.forEach((group) => {
                group.voucherTypes.forEach((vt: any) => {
                    const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                    const isExpanded = expandedVouchers.includes(vKey);
                    const filteredStaff = getFilteredStaffList(vt.staff || [], selectedStaffFilters, roleColumns, qtyType);
                    const hasStaff = filteredStaff.length > 0;

                    // 1. Voucher Type (Main Row)
                    const vtRow: any = {};
                    vtRow["Category"] = category.name;
                    vtRow["Groups"] = group.groupName;
                    vtRow["Group Kgs"] = group.groupKgs;
                    vtRow["Voucher Type"] = vt.name;
                    vtRow["Voucher Kgs"] = vt.baseKgs;
                    vtRow["Total Qty"] = vt.roleSums["Total_Qty"] || 0;
                    vtRow["Total Count"] = vt.roleSums["Total_Count"] || 0;
                    vtRow["Staff Name / Inv No"] = "-";
                    enabledRoles.forEach(col => {
                        const val = vt.roleSums[col.key] || 0;
                        vtRow[getColumnLabel(col)] = val;
                    });
                    rows.push(vtRow);

                    // 2. Staff Rows (if expanded)
                    if (isExpanded && hasStaff) {
                        filteredStaff.forEach((sc: any) => {
                            const staffKey = `${vKey}_${sc.Emp_Id || 'unassigned'}`;
                            const isStaffExpanded = expandedStaff.includes(staffKey);
                            const invoicesList = getFilteredStaffByCategories(invoiceData[staffKey] || [], roleColumns, qtyType);
                            const hasInvoices = invoicesList.length > 0;
                            const staffRow: any = {};
                            staffRow["Category"] = "";
                            staffRow["Groups"] = "";
                            staffRow["Group Kgs"] = "";
                            staffRow["Voucher Type"] = "";
                            staffRow["Voucher Kgs"] = "";
                            staffRow["Total Qty"] = getStaffCellValue(sc, "Total_Qty", "qty", qtyType, roleColumns);
                            staffRow["Total Count"] = getStaffCellValue(sc, "Total_Count", "count", qtyType, roleColumns);
                            staffRow["Staff Name / Inv No"] = sc.Cost_Center_Name || "Unassigned";
                            enabledRoles.forEach(col => {
                                const val = getStaffCellValue(sc, col.key, col.metric || "qty", qtyType, roleColumns);
                                staffRow[getColumnLabel(col)] = val;
                            });
                            rows.push(staffRow);

                            // 3. Invoice Rows (if staff is expanded)
                            if (isStaffExpanded && hasInvoices) {
                                invoicesList.forEach((inv: any) => {
                                    const invRow: any = {};
                                    invRow["Category"] = "";
                                    invRow["Groups"] = "";
                                    invRow["Group Kgs"] = "";
                                    invRow["Voucher Type"] = "";
                                    invRow["Voucher Kgs"] = "";
                                    invRow["Total Qty"] = getInvoiceCellValue(inv, "Total_Qty", "qty", qtyType, roleColumns);
                                    invRow["Total Count"] = getInvoiceCellValue(inv, "Total_Count", "count", qtyType, roleColumns);
                                    invRow["Staff Name / Inv No"] = inv.Inv_No;
                                    enabledRoles.forEach(col => {
                                        const val = getInvoiceCellValue(inv, col.key, col.metric || "qty", qtyType, roleColumns);
                                        invRow[getColumnLabel(col)] = val;
                                    });
                                    rows.push(invRow);
                                });
                            }
                        });
                    }
                });
            });
        });

        // 4. Grand Total Row
        const grandRow: any = {};
        grandRow["Category"] = "Total";
        grandRow["Groups"] = "Grand Total";
        grandRow["Group Kgs"] = grandTotals.totalKgs;
        grandRow["Voucher Type"] = "";
        grandRow["Voucher Kgs"] = "";
        grandRow["Total Qty"] = grandTotals.roleTotals["Total_Qty"] || 0;
        grandRow["Total Count"] = grandTotals.roleTotals["Total_Count"] || 0;
        grandRow["Staff Name / Inv No"] = "";
        enabledRoles.forEach(col => {
            grandRow[getColumnLabel(col)] = grandTotals.roleTotals[col.key] || 0;
        });
        rows.push(grandRow);

        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Overall Staff Report");
        XLSX.writeFile(workbook, `Overall Staff Report_${dayjs().format("DDMMYYYY")}.xlsx`);
        toast.success("Excel Exported ✅");
    };

    const handleExportPDF = () => {
        const cols = [
            "Category",
            "Groups",
            "Group Kgs",
            "Voucher Type",
            "Voucher Kgs",
            "Total Qty",
            "Total Count",
            "Staff Name / Inv No"
        ];

        const enabledRoles = roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order);
        enabledRoles.forEach(col => {
            cols.push(getColumnLabel(col));
        });

        const rows: any[] = [];

        tableCategories.forEach((category) => {
            category.groups.forEach((group) => {
                group.voucherTypes.forEach((vt: any) => {
                    const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                    const isExpanded = expandedVouchers.includes(vKey);
                    const filteredStaff = getFilteredStaffList(vt.staff || [], selectedStaffFilters, roleColumns, qtyType);
                    const hasStaff = filteredStaff.length > 0;

                    // 1. Voucher Type (Main Row)
                    const vtRowData = [
                        category.name,
                        group.groupName,
                        group.groupKgs.toLocaleString(),
                        vt.name,
                        vt.baseKgs.toLocaleString(),
                        (vt.roleSums["Total_Qty"] || 0).toLocaleString(),
                        (vt.roleSums["Total_Count"] || 0).toLocaleString(),
                        "-"
                    ];
                    enabledRoles.forEach(col => {
                        const val = vt.roleSums[col.key] || 0;
                        vtRowData.push(safeLocaleString(val));
                    });
                    rows.push(vtRowData);

                    // 2. Staff Rows (if expanded)
                    if (isExpanded && hasStaff) {
                        filteredStaff.forEach((sc: any) => {
                            const staffKey = `${vKey}_${sc.Emp_Id || 'unassigned'}`;
                            const isStaffExpanded = expandedStaff.includes(staffKey);
                            const invoicesList = getFilteredStaffByCategories(invoiceData[staffKey] || [], roleColumns, qtyType);
                            const hasInvoices = invoicesList.length > 0;

                            const staffRowData = [
                                "",
                                "",
                                "",
                                "",
                                "",
                                getStaffCellValue(sc, "Total_Qty", "qty", qtyType, roleColumns).toLocaleString(),
                                getStaffCellValue(sc, "Total_Count", "count", qtyType, roleColumns).toLocaleString(),
                                sc.Cost_Center_Name || "Unassigned"
                            ];
                            enabledRoles.forEach(col => {
                                const val = getStaffCellValue(sc, col.key, col.metric || "qty", qtyType, roleColumns);
                                staffRowData.push(safeLocaleString(val));
                            });
                            rows.push(staffRowData);

                            // 3. Invoice Rows (if staff is expanded)
                            if (isStaffExpanded && hasInvoices) {
                                invoicesList.forEach((inv: any) => {
                                    const invRowData = [
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        getInvoiceCellValue(inv, "Total_Qty", "qty", qtyType, roleColumns).toLocaleString(),
                                        getInvoiceCellValue(inv, "Total_Count", "count", qtyType, roleColumns).toLocaleString(),
                                        inv.Inv_No
                                    ];
                                    enabledRoles.forEach(col => {
                                        const val = getInvoiceCellValue(inv, col.key, col.metric || "qty", qtyType, roleColumns);
                                        invRowData.push(safeLocaleString(val));
                                    });
                                    rows.push(invRowData);
                                });
                            }
                        });
                    }
                });
            });
        });

        // 4. Grand Total Row
        const grandRowData = [
            "Total",
            "Grand Total",
            grandTotals.totalKgs.toLocaleString(),
            "",
            "",
            (grandTotals.roleTotals["Total_Qty"] || 0).toLocaleString(),
            (grandTotals.roleTotals["Total_Count"] || 0).toLocaleString(),
            ""
        ];
        enabledRoles.forEach(col => {
            const val = grandTotals.roleTotals[col.key] || 0;
            grandRowData.push(safeLocaleString(val));
        });
        rows.push(grandRowData);

        const doc = new jsPDF("l", "mm", "a4");
        doc.text("Overall Staff Report", 14, 10);
        autoTable(doc, {
            startY: 15,
            head: [cols],
            body: rows,
            styles: { fontSize: 7 },
            headStyles: { fillColor: [30, 58, 138] },
        });

        doc.save(`Overall Staff Report_${dayjs().format("DDMMYYYY")}.pdf`);
        toast.success("PDF Exported ✅");
    };

    return (
        <>
            <PageHeader
                parentReportName="Overall Staff Report"
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
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
                fromDate={tempFromDate}
                onFromDateChange={setTempFromDate}
                toDate={tempToDate}
                onToDateChange={setTempToDate}
                onApply={() => {
                    setFromDate(tempFromDate);
                    setToDate(tempToDate);
                    setDrawerOpen(false);
                }}
            >
                {/* Qty / Act Qty Selection */}
                <FormControl sx={{ mb: 2, display: "block" }}>
                    <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                        Quantity Type
                    </FormLabel>
                    <RadioGroup
                        value={qtyType}
                        onChange={(e) => handleQtyTypeChange(e.target.value as "Qty" | "Act_Qty")}
                    >
                        <FormControlLabel
                            value="Qty"
                            control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                            label={<Typography sx={{ fontSize: "0.825rem" }}>Qty</Typography>}
                        />
                        <FormControlLabel
                            value="Act_Qty"
                            control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                            label={<Typography sx={{ fontSize: "0.825rem" }}>Act Qty</Typography>}
                        />
                    </RadioGroup>
                </FormControl>
            </ReportFilterDrawer>

            <AppLayout fullWidth>
                <Box px={2} pb={1} pt={1}>
                    {(loading || preloading) ? (
                        <Box display="flex" flexDirection="column" alignItems="center" justifyContent="center" height="60vh" gap={2}>
                            <CircularProgress size={40} />
                            <Typography variant="body1" color="text.secondary" fontWeight={500}>
                                Loading report data and template settings...
                            </Typography>
                        </Box>
                    ) : (
                        <TableContainer component={Paper} elevation={1} sx={{ borderRadius: 2, border: "1px solid #cbd5e1", overflow: "auto", maxHeight: "calc(100vh - 90px)" }}>
                        <Table size="medium" stickyHeader>
                            <TableHead>
                                <TableRow sx={{ bgcolor: "#1E3A8A" }}>
                                    <TableCell sx={headerCellSx} align="center">
                                        Category
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Groups
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Kgs
                                    </TableCell>
                                    <TableCell sx={headerCellSx}>
                                        Voucher Type
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Kgs
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Total Qty
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Total Count
                                    </TableCell>
                                    <TableCell
                                        sx={{
                                            ...headerCellSx,
                                            cursor: "pointer",
                                            "&:hover": { bgcolor: "#1b3580" }
                                        }}
                                        onClick={(e) => setStaffFilterAnchor(e.currentTarget)}
                                    >
                                        STAFF NAME
                                    </TableCell>

                                    {roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order).map(col => (
                                        <TableCell key={col.key} sx={headerCellSx} align="center">
                                            {getColumnLabel(col)}
                                        </TableCell>
                                    ))}
                                </TableRow>
                                <TableRow sx={{ borderBottom: "2px solid #1e3a8a", bgcolor: "#bfdbfe" }}>
                                    <TableCell colSpan={3} sx={totalCellSx} align="center">
                                        Total
                                    </TableCell>
                                    <TableCell sx={{ ...totalCellSx, borderRight: "1px solid #e2e8f0" }}>
                                        Grand Total
                                    </TableCell>
                                    <TableCell sx={{ ...totalCellSx, borderRight: "1px solid #e2e8f0" }} align="center">
                                        {grandTotals.totalKgs.toLocaleString()}
                                    </TableCell>
                                    <TableCell sx={{ ...totalCellSx, borderRight: "1px solid #e2e8f0" }} align="center">
                                        {grandTotals.roleTotals["Total_Qty"].toLocaleString()}
                                    </TableCell>
                                    <TableCell sx={{ ...totalCellSx, borderRight: "1px solid #e2e8f0" }} align="center">
                                        {grandTotals.roleTotals["Total_Count"].toLocaleString()}
                                    </TableCell>
                                    <TableCell sx={{ ...totalCellSx, borderRight: "1px solid #93c5fd" }} />
                                    {roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order).map(col => (
                                        <TableCell key={col.key} sx={{ ...totalCellSx, borderRight: "1px solid #e2e8f0" }} align="center">
                                            {grandTotals.roleTotals[col.key].toLocaleString()}
                                        </TableCell>
                                    ))}
                                </TableRow>
                            </TableHead>
                            <TableBody>
                                {tableCategories.length > 0 ? (
                                    tableCategories.map((category) => {
                                        let categorySpan = 0;
                                        category.groups.forEach((group) => {
                                            group.voucherTypes.forEach((vt: any) => {
                                                categorySpan += 1;
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                if (expandedVouchers.includes(vKey)) {
                                                    const staffList = staffData[vKey] || [];
                                                    const filteredStaffList = getFilteredStaffList(staffList, selectedStaffFilters, roleColumns, qtyType);
                                                    categorySpan += filteredStaffList.length;
                                                    filteredStaffList.forEach((sc: any) => {
                                                        const staffKey = `${vKey}_${sc.Emp_Id || 'unassigned'}`;
                                                        if (expandedStaff.includes(staffKey)) {
                                                            const invoicesList = getFilteredStaffByCategories(invoiceData[staffKey] || [], roleColumns, qtyType);
                                                            categorySpan += invoicesList.length;
                                                        }
                                                    });
                                                }
                                            });
                                        });

                                        let isFirstCategoryRow = true;

                                        return category.groups.map((group) => {
                                            let groupSpan = 0;
                                            group.voucherTypes.forEach((vt: any) => {
                                                groupSpan += 1;
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                if (expandedVouchers.includes(vKey)) {
                                                    const staffList = staffData[vKey] || [];
                                                    const filteredStaffList = getFilteredStaffList(staffList, selectedStaffFilters, roleColumns, qtyType);
                                                    groupSpan += filteredStaffList.length;
                                                    filteredStaffList.forEach((sc: any) => {
                                                        const staffKey = `${vKey}_${sc.Emp_Id || 'unassigned'}`;
                                                        if (expandedStaff.includes(staffKey)) {
                                                            const invoicesList = getFilteredStaffByCategories(invoiceData[staffKey] || [], roleColumns, qtyType);
                                                            groupSpan += invoicesList.length;
                                                        }
                                                    });
                                                }
                                            });

                                            let isFirstGroupRow = true;

                                            return group.voucherTypes.map((vt: any) => {
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                const isExpanded = expandedVouchers.includes(vKey);
                                                const filteredStaff = getFilteredStaffList(vt.staff || [], selectedStaffFilters, roleColumns, qtyType);
                                                const hasStaff = filteredStaff.length > 0;

                                                const mainRow = (
                                                    <TableRow key={vt.name} sx={{ bgcolor: "#f8fafc", "&:hover": { bgcolor: "#f1f5f9" } }}>
                                                        {isFirstCategoryRow && (
                                                            <TableCell
                                                                rowSpan={categorySpan}
                                                                sx={{
                                                                    fontWeight: 800,
                                                                    bgcolor: "#e0f2fe",
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

                                                        {isFirstGroupRow && (
                                                            <TableCell
                                                                rowSpan={groupSpan}
                                                                sx={{
                                                                    fontWeight: 750,
                                                                    bgcolor: "#f0f9ff",
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

                                                        {isFirstGroupRow && (
                                                            <TableCell
                                                                rowSpan={groupSpan}
                                                                sx={{
                                                                    fontWeight: 750,
                                                                    bgcolor: "#f0f9ff",
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
                                                            onClick={() => handleToggleExpandVoucher(category.name, group.groupName, vt.name)}
                                                        >
                                                            <Box display="flex" alignItems="center" gap={0.5}>
                                                                {isExpanded ? <KeyboardArrowUpIcon fontSize="small" /> : <KeyboardArrowDownIcon fontSize="small" />}
                                                                {vt.name}
                                                            </Box>
                                                        </TableCell>

                                                        <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #cbd5e1", fontSize: "0.825rem", py: 1 }} align="center">
                                                            {vt.baseKgs.toLocaleString()}
                                                        </TableCell>

                                                        <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #cbd5e1", fontSize: "0.825rem", py: 1 }} align="center">
                                                            {(vt.roleSums["Total_Qty"] || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #cbd5e1", fontSize: "0.825rem", py: 1 }} align="center">
                                                            {(vt.roleSums["Total_Count"] || 0).toLocaleString()}
                                                        </TableCell>

                                                        <TableCell sx={{ fontStyle: "italic", color: "text.secondary", borderRight: "1px solid #cbd5e1", py: 1 }} align="center">
                                                            -
                                                        </TableCell>

                                                        {roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order).map(col => {
                                                            const value = vt.roleSums[col.key] || 0;
                                                            return (
                                                                <TableCell key={col.key} sx={{ fontWeight: 800, borderRight: "1px solid #cbd5e1", py: 1, bgcolor: "#f1f5f9" }} align="center">
                                                                    {value > 0 ? value.toLocaleString() : "0"}
                                                                </TableCell>
                                                            );
                                                        })}
                                                    </TableRow>
                                                );

                                                isFirstCategoryRow = false;
                                                isFirstGroupRow = false;

                                                const staffRows = isExpanded && hasStaff ? (
                                                    filteredStaff.map((sc: any) => {
                                                        const staffKey = `${vKey}_${sc.Emp_Id || 'unassigned'}`;
                                                        const isStaffExpanded = expandedStaff.includes(staffKey);
                                                        const invoicesList = getFilteredStaffByCategories(invoiceData[staffKey] || [], roleColumns, qtyType);
                                                        const hasInvoices = invoicesList.length > 0;

                                                        const staffRow = (
                                                            <TableRow key={sc.Cost_Center_Id || sc.Cost_Center_Name} sx={{ "&:hover": { bgcolor: "#f8fafc" } }}>
                                                                <TableCell sx={{ borderRight: "1px solid #e2e8f0", py: 0.8 }} />
                                                                <TableCell sx={{ borderRight: "1px solid #e2e8f0", py: 0.8 }} align="center" />

                                                                <TableCell sx={{ color: "#475569", borderRight: "1px solid #e2e8f0", py: 0.8, fontSize: "0.8rem" }} align="center">
                                                                    {getStaffCellValue(sc, "Total_Qty", "qty", qtyType, roleColumns).toLocaleString()}
                                                                </TableCell>
                                                                <TableCell sx={{ color: "#475569", borderRight: "1px solid #e2e8f0", py: 0.8, fontSize: "0.8rem" }} align="center">
                                                                    {getStaffCellValue(sc, "Total_Count", "count", qtyType, roleColumns).toLocaleString()}
                                                                </TableCell>

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
                                                                    onClick={() => handleToggleExpandStaff(category.name, group.groupName, vt.name, sc.Emp_Id, sc.Cost_Center_Name)}
                                                                >
                                                                    <Box display="flex" alignItems="center" gap={0.5}>
                                                                        {isStaffExpanded ? (
                                                                            <KeyboardArrowUpIcon sx={{ fontSize: "0.9rem" }} />
                                                                        ) : (
                                                                            <KeyboardArrowDownIcon sx={{ fontSize: "0.9rem" }} />
                                                                        )}
                                                                        {sc.Cost_Center_Name}
                                                                    </Box>
                                                                </TableCell>

                                                                {roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order).map(col => {
                                                                    const metric: "qty" | "count" = col.metric || "qty";
                                                                    const roleVal = getStaffCellValue(sc, col.key, metric, qtyType, roleColumns);
                                                                    return (
                                                                        <TableCell key={col.key} sx={{ color: "#475569", borderRight: "1px solid #e2e8f0", py: 0.8, fontSize: "0.8rem" }} align="center">
                                                                            {roleVal > 0 ? roleVal.toLocaleString() : "0"}
                                                                        </TableCell>
                                                                    );
                                                                })}
                                                            </TableRow>
                                                        );

                                                        const invoiceRows = isStaffExpanded && hasInvoices ? (
                                                            invoicesList.map((inv: any, idx: number) => (
                                                                <TableRow key={`${staffKey}_${inv.Inv_No}_${idx}`} sx={{ bgcolor: "#ffffff", "&:hover": { bgcolor: "#f8fafc" } }}>
                                                                    <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.6 }} />
                                                                    <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.6 }} align="center" />

                                                                    <TableCell sx={{ color: "#475569", borderRight: "1px solid #cbd5e1", py: 0.6, fontSize: "0.8rem" }} align="center">
                                                                        {getInvoiceCellValue(inv, "Total_Qty", "qty", qtyType, roleColumns).toLocaleString()}
                                                                    </TableCell>
                                                                    <TableCell sx={{ color: "#475569", borderRight: "1px solid #cbd5e1", py: 0.6, fontSize: "0.8rem" }} align="center">
                                                                        {getInvoiceCellValue(inv, "Total_Count", "count", qtyType, roleColumns).toLocaleString()}
                                                                    </TableCell>

                                                                    <TableCell
                                                                        sx={{
                                                                            color: "#334155",
                                                                            borderRight: "1px solid #cbd5e1",
                                                                            py: 0.6,
                                                                            fontSize: "0.8rem",
                                                                            pl: 6
                                                                        }}
                                                                    >
                                                                        {inv.Inv_No}
                                                                    </TableCell>

                                                                    {roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").sort((a, b) => a.order - b.order).map(col => {
                                                                        const metric: "qty" | "count" = col.metric || "qty";
                                                                        const val = getInvoiceCellValue(inv, col.key, metric, qtyType, roleColumns);
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
                                                                                {val !== undefined && val !== null ? val.toLocaleString() : "0"}
                                                                            </TableCell>
                                                                        );
                                                                    })}
                                                                </TableRow>
                                                            ))
                                                        ) : null;

                                                        return (
                                                            <React.Fragment key={sc.Emp_Id || sc.Cost_Center_Name}>
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
                                        <TableCell colSpan={8 + roleColumns.filter(c => c.enabled && c.key !== "Total_Qty" && c.key !== "Total_Count").length} align="center" sx={{ py: 6 }}>
                                            No records found.
                                        </TableCell>
                                    </TableRow>
                                )}
                            </TableBody>
                        </Table>
                        </TableContainer>
                    )}
                </Box>
            </AppLayout>

            {/* Staff Filter Popover */}
            <Popover
                open={Boolean(staffFilterAnchor)}
                anchorEl={staffFilterAnchor}
                onClose={() => {
                    setStaffFilterAnchor(null);
                    setStaffSearchQuery("");
                }}
                anchorOrigin={{
                    vertical: "bottom",
                    horizontal: "left",
                }}
                transformOrigin={{
                    vertical: "top",
                    horizontal: "left",
                }}
            >
                <Box p={2} width={250} display="flex" flexDirection="column" gap={1}>
                    <TextField
                        size="small"
                        placeholder="Search staff..."
                        value={staffSearchQuery}
                        onChange={(e) => setStaffSearchQuery(e.target.value)}
                        fullWidth
                    />

                    <Box sx={{ maxHeight: 250, overflowY: "auto", my: 1, display: "flex", flexDirection: "column", gap: 0.5 }}>
                        {filteredGridStaffNames.length === 0 ? (
                            <Typography variant="caption" color="text.secondary" p={1}>
                                No staff found
                            </Typography>
                        ) : (
                            <>
                                <FormControlLabel
                                    control={
                                        <Checkbox
                                            size="small"
                                            checked={selectedStaffFilters.length === 0}
                                            onChange={handleToggleSelectAll}
                                        />
                                    }
                                    label="All"
                                    sx={{ margin: 0, "& .MuiFormControlLabel-label": { fontSize: "0.8rem", fontWeight: 700 } }}
                                />
                                {filteredGridStaffNames.map(name => {
                                    const isChecked = selectedStaffFilters.includes(name);
                                    return (
                                        <FormControlLabel
                                            key={name}
                                            control={
                                                <Checkbox
                                                    size="small"
                                                    checked={isChecked}
                                                    onChange={() => handleToggleStaff(name)}
                                                />
                                            }
                                            label={name}
                                            sx={{ margin: 0, "& .MuiFormControlLabel-label": { fontSize: "0.8rem" } }}
                                        />
                                    );
                                })}
                            </>
                        )}
                    </Box>
                </Box>
            </Popover>

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

                    {preloading ? (
                        <Box display="flex" flexDirection="column" alignItems="center" justifyContent="center" py={4} gap={1.5}>
                            <CircularProgress size={30} />
                            <Typography variant="caption" color="text.secondary">
                                Loading Column Config...
                            </Typography>
                        </Box>
                    ) : (
                        <DndContext
                            sensors={sensors}
                            collisionDetection={closestCenter}
                            onDragEnd={handleDragEnd}
                        >
                            <SortableContext
                                items={roleColumns.filter(c => c.key !== "Total_Qty" && c.key !== "Total_Count").map(c => c.key)}
                                strategy={verticalListSortingStrategy}
                            >
                                <Box display="flex" flexDirection="column" gap={0.5}>
                                    {[...roleColumns]
                                        .filter(c => c.key !== "Total_Qty" && c.key !== "Total_Count")
                                        .sort((a, b) => a.order - b.order)
                                        .map((col) => (
                                            <SortableColumnRow
                                                key={col.key}
                                                column={col}
                                                onToggle={handleToggleColumn}
                                                onChangeMetric={handleChangeMetric}
                                            />
                                        ))}
                                </Box>
                            </SortableContext>
                        </DndContext>
                    )}
                </DialogContent>
            </Dialog>

            {/* Group Dialog */}
            <Dialog
                open={groupCreateOpen}
                onClose={handleCloseGroupDialog}
                maxWidth="sm"
                fullWidth
            >
                <DialogTitle sx={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    Group Management
                    <IconButton size="small" onClick={handleCloseGroupDialog}>
                        <CloseIcon />
                    </IconButton>
                </DialogTitle>
                <DialogContent dividers>
                    <Autocomplete
                        freeSolo
                        options={groups.map(g => g.name)}
                        value={newGroupName}
                        onChange={(_, newValue) => {
                            if (newValue) {
                                const grp = groups.find(g => g.name === newValue);
                                if (grp) {
                                    setSelectedGroupToEdit(grp.name);
                                    setNewGroupName(grp.name);
                                    setNewGroupCategory(grp.parentCategory);
                                    setNewGroupVouchers(grp.voucherTypes);
                                    return;
                                }
                            }
                            setSelectedGroupToEdit("");
                            setNewGroupName(newValue || "");
                        }}
                        onInputChange={(_, newInputValue) => {
                            setNewGroupName(newInputValue);
                            const grp = groups.find(g => g.name.toLowerCase() === newInputValue.trim().toLowerCase());
                            if (grp) {
                                setSelectedGroupToEdit(grp.name);
                                setNewGroupCategory(grp.parentCategory);
                                setNewGroupVouchers(grp.voucherTypes);
                            } else {
                                setSelectedGroupToEdit("");
                            }
                        }}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                label="Group Name"
                                placeholder="Select existing group or type new..."
                                size="small"
                                sx={{ mb: 2.5 }}
                                fullWidth
                            />
                        )}
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

                    <Autocomplete
                        multiple
                        size="small"
                        options={allVoucherNames}
                        value={newGroupVouchers}
                        onChange={(_, newValue) => setNewGroupVouchers(newValue)}
                        disableCloseOnSelect
                        getOptionLabel={(option) => option}
                        renderOption={(props, option, { selected }) => {
                            const { key, ...rest } = props;
                            return (
                                <li key={key} {...rest}>
                                    <Checkbox
                                        checked={selected}
                                        size="small"
                                        style={{ marginRight: 8 }}
                                    />
                                    {option}
                                </li>
                            );
                        }}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                label="Voucher Types"
                                placeholder="Select voucher types..."
                            />
                        )}
                        fullWidth
                    />
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleCloseGroupDialog}>Cancel</Button>
                    <Button
                        onClick={handleCreateGroup}
                        variant="contained"
                        color="primary"
                    >
                        {selectedGroupToEdit ? "Save Changes" : "Create Group"}
                    </Button>
                </DialogActions>
            </Dialog>

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