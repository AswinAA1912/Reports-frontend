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
    Switch,
    FormControlLabel,
    FormControl,
    Menu,
    CircularProgress,
    RadioGroup,
    Radio,
    FormLabel,
    Chip
} from "@mui/material";
import SettingsIcon from "@mui/icons-material/Settings";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowUpIcon from "@mui/icons-material/KeyboardArrowUp";
import CloseIcon from "@mui/icons-material/Close";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import FilterAltIcon from "@mui/icons-material/FilterAlt";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import { toast } from "react-toastify";
import { SettingsService } from "../../services/reportSettings.services";
import { employeeReportGroupService, VoucherType, GodownMaster } from "../../services/staffBasedReport.services";
import { useAuth } from "../../auth/authContext";
import { DndContext, closestCenter, PointerSensor, useSensor, useSensors } from "@dnd-kit/core";
import { SortableContext, useSortable, verticalListSortingStrategy, arrayMove } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";

// Column config type for dynamic columns
export interface ColumnConfig {
    key: string;
    label: string;
    enabled: boolean;
    order: number;
}

// Default columns for the 2nd API (invoicewithitems)
// Brand and Item_Name_Modified are enabled by default as requested
const DEFAULT_ITEM_COLUMNS: ColumnConfig[] = [
    { key: "Brand", label: "Brand", enabled: true, order: 0 },
    { key: "Item_Name_Modified", label: "Item_Name_Modified", enabled: true, order: 1 },
    { key: "Stock_Item", label: "Stock_Item", enabled: false, order: 2 },
    { key: "Group_ST", label: "Group_ST", enabled: false, order: 3 },
    { key: "Bag", label: "Bag", enabled: false, order: 4 },
    { key: "Stock_Group", label: "Stock_Group", enabled: false, order: 5 },
    { key: "S_Sub_Group_1", label: "S_Sub_Group_1", enabled: false, order: 6 },
    { key: "Grade_Item_Group", label: "Grade_Item_Group", enabled: false, order: 7 },
    { key: "POS_Group", label: "POS_Group", enabled: false, order: 8 },
    { key: "POS_Item_Name", label: "POS_Item_Name", enabled: false, order: 9 },
    { key: "Inv_No", label: "Inv_No", enabled: false, order: 10 },
    { key: "Voucher_Date", label: "Voucher_Date", enabled: false, order: 11 },
    { key: "Cost_Center_Name", label: "Cost_Center_Name", enabled: false, order: 12 },
    { key: "Cost_Category", label: "Cost_Category", enabled: false, order: 13 },
    { key: "Brokerage", label: "Brokerage", enabled: false, order: 14 },
    { key: "Coolie", label: "Coolie", enabled: false, order: 15 },
    { key: "Stock_Date_Added", label: "Stock_Date_Added", enabled: false, order: 16 },
];

const EXCLUDED_ITEM_KEYS = new Set([
    "Voucher_Ref_Id",
    "Ref_Id",
    "Bill_Qty",
    "Act_Qty",
    "Total_Qty",
    "Total_Act_Qty",
    "Voucher_Type",
    "Voucher_Type_Id",
    "Group_Name",
    "Overall_GroupName",
    "Cost_Center_Id",
    "Active",
]);

const discoverItemColumns = (items: any[], currentCols: ColumnConfig[]): ColumnConfig[] => {
    if (!items || items.length === 0) return currentCols;

    const discoveredKeys = new Set<string>();
    items.forEach((row) => {
        Object.keys(row).forEach((key) => {
            if (EXCLUDED_ITEM_KEYS.has(key)) return;
            discoveredKeys.add(key);
        });
    });

    const newCols = [...currentCols];
    discoveredKeys.forEach((key) => {
        const exists = newCols.some((c) => c.key.toUpperCase() === key.toUpperCase());
        if (!exists) {
            newCols.push({
                key: key,
                label: key,
                enabled: false,
                order: newCols.length,
            });
        }
    });

    return newCols.map((c, index) => ({ ...c, order: index }));
};

interface SortableColumnRowProps {
    column: ColumnConfig;
    onToggle: (key: string) => void;
    hasActiveFilter?: boolean;
}

const SortableColumnRow = ({ column, onToggle, hasActiveFilter }: SortableColumnRowProps) => {
    const {
        attributes,
        listeners,
        setNodeRef,
        transform,
        transition,
    } = useSortable({ id: column.key });

    const style = {
        transform: CSS.Transform.toString(transform),
        transition,
    };

    return (
        <Box
            ref={setNodeRef}
            style={style}
            display="flex"
            alignItems="center"
            gap={1}
            mb={0.5}
            px={0.5}
        >
            <Box
                {...listeners}
                {...attributes}
                sx={{
                    display: "flex",
                    alignItems: "center",
                    cursor: "grab",
                    color: "text.secondary",
                    p: 0.5
                }}
            >
                <DragIndicatorIcon fontSize="small" />
            </Box>

            <Box display="flex" alignItems="center" gap={1} sx={{ flex: 1 }}>
                <Typography fontSize="0.75rem">
                    {column.label}
                </Typography>
                {Boolean(hasActiveFilter) && (
                    <Tooltip title="Header filter enabled">
                        <FilterAltIcon fontSize="small" color="primary" sx={{ fontSize: "0.85rem" }} />
                    </Tooltip>
                )}
            </Box>

            <Switch
                size="medium"
                checked={column.enabled}
                onChange={() => onToggle(column.key)}
                sx={{
                    "& .MuiSwitch-switchBase.Mui-checked": {
                        color: "#1E3A8A",
                    },
                    "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": {
                        backgroundColor: "#b5b9c4",
                    },
                    "& .MuiSwitch-track": {
                        backgroundColor: "#CBD5E1",
                    },
                }}
            />
        </Box>
    );
};

// Group Mapping Config
interface GroupConfig {
    name: string;
    parentCategory: "INWARDS" | "OUTWARDS" | "ADJUSTMENTS";
    voucherTypes: string[];
}

const OverallItemwise: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const { user } = useAuth();
    const companyId = user?.companyId || "default";

    // Date & UI States
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [tempFromDate, setTempFromDate] = useState(today);
    const [tempToDate, setTempToDate] = useState(today);
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [loading, setLoading] = useState(false);

    // Godown Master State & Filter
    const [godowns, setGodowns] = useState<GodownMaster[]>([]);
    const [selectedGodownIds, setSelectedGodownIds] = useState<number[]>([]);
    const [tempSelectedGodownIds, setTempSelectedGodownIds] = useState<number[]>([]);

    // Quantity Type: "Qty" vs "Act_Qty"
    const [qtyType, setQtyType] = useState<"Qty" | "Act_Qty">("Qty");

    useEffect(() => {
        if (companyId) {
            const stored = localStorage.getItem(`group_wise_item_qty_type_${companyId}`);
            setQtyType(stored === "Act_Qty" ? "Act_Qty" : "Qty");
        }
    }, [companyId]);

    const handleQtyTypeChange = (newType: "Qty" | "Act_Qty") => {
        setQtyType(newType);
        if (companyId) {
            localStorage.setItem(`group_wise_item_qty_type_${companyId}`, newType);
        }
    };

    useEffect(() => {
        if (drawerOpen) {
            setTempFromDate(fromDate);
            setTempToDate(toDate);
            setTempSelectedGodownIds(selectedGodownIds);
        }
    }, [drawerOpen, fromDate, toDate, selectedGodownIds]);

    // Live Report Data State (from 1st API: overallStaffCategorywise)
    const [reportData, setReportData] = useState<any[]>([]);

    // Cached Item Data (from 2nd API: invoicewithitems) keyed by voucherKey
    const [itemData, setItemData] = useState<Record<string, any[]>>({});
    const [loadingItems, setLoadingItems] = useState<Record<string, boolean>>({});

    // Column Config State for 2nd API columns
    const [itemColumns, setItemColumns] = useState<ColumnConfig[]>(DEFAULT_ITEM_COLUMNS);

    // Enabled columns sorted by order
    const enabledItemColumns = useMemo(() => {
        return itemColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order);
    }, [itemColumns]);

    // Template states
    const [saveDialogOpen, setSaveDialogOpen] = useState(false);
    const [exportModalOpen, setExportModalOpen] = useState(false);
    const [exportType, setExportType] = useState<"excel" | "pdf">("excel");
    const [exportDetails, setExportDetails] = useState(false);
    const [reportName, setReportName] = useState("");
    const parentReportName = "GROUP WISE ITEM";
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);

    // 2nd API Column Filters & Header Filter Menu
    const [columnFilters, setColumnFilters] = useState<Record<string, string[] | undefined>>({});
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);

    // Column Settings Popover/Dialog Anchor
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);

    // Expanded Vouchers state
    const [expandedVouchers, setExpandedVouchers] = useState<string[]>([]);

    // Grouping Config & Master Voucher Types
    const [groups, setGroups] = useState<GroupConfig[]>([]);
    const [voucherTypes, setVoucherTypes] = useState<VoucherType[]>([]);

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
        if (!over || active.id === over.id) return;

        setItemColumns((prev) => {
            const enabledCols = prev.filter(c => c.enabled).sort((a, b) => a.order - b.order);
            const oldIndex = enabledCols.findIndex(c => c.key === active.id);
            const newIndex = enabledCols.findIndex(c => c.key === over.id);
            if (oldIndex === -1 || newIndex === -1) return prev;
            const reordered = arrayMove(enabledCols, oldIndex, newIndex);

            return prev.map((col) => {
                const idx = reordered.findIndex(r => r.key === col.key);
                return idx !== -1 ? { ...col, order: idx } : col;
            });
        });
    };

    // Toggle column visibility
    const handleToggleColumn = (key: string) => {
        setItemColumns(p => p.map(col => col.key === key ? { ...col, enabled: !col.enabled } : col));
    };

    // Fetch voucher types, godowns, and groups concurrently on mount
    useEffect(() => {
        const fetchMetadata = async () => {
            try {
                const [vtRes, gRes, ergRes] = await Promise.allSettled([
                    employeeReportGroupService.getVoucherTypes(),
                    employeeReportGroupService.getGodownsWithVoucherTypes().catch(() => employeeReportGroupService.getGodowns()),
                    employeeReportGroupService.getEmployeeReportGroups()
                ]);

                if (vtRes.status === "fulfilled" && vtRes.value?.data?.success) {
                    setVoucherTypes(vtRes.value.data.data || []);
                }

                if (gRes.status === "fulfilled" && gRes.value?.data?.success) {
                    const gData = gRes.value.data.data || [];
                    const filteredG = gData.filter((g: any) => g.Godown_Name && g.Godown_Name.trim() !== "");
                    setGodowns(filteredG);
                }

                if (ergRes.status === "fulfilled" && ergRes.value?.data?.success) {
                    const rows = ergRes.value.data.data || [];
                    const groupsMap: Record<string, {
                        name: string;
                        parentCategory: "INWARDS" | "OUTWARDS" | "ADJUSTMENTS";
                        voucherTypes: string[];
                    }> = {};

                    rows.forEach((row: any) => {
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

    // Fetch 1st API: overallStaffCategorywise
    const fetchReportData = async () => {
        setLoading(true);
        try {
            const res = await employeeReportGroupService.getOverallStaffCategorywise({
                Fromdate: fromDate,
                Todate: toDate
            });
            if (res.data?.success) {
                const data = res.data.data || [];
                setReportData(data);
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
        fetchReportData();
        setExpandedVouchers([]);
        setItemData({});
    }, [fromDate, toDate]);

    // Set of voucher type names linked to the selected godown(s)
    const linkedVoucherNames = useMemo(() => {
        if (!selectedGodownIds || selectedGodownIds.length === 0) return null;
        const godownIdSet = new Set(selectedGodownIds.map(Number));
        const set = new Set<string>();
        voucherTypes.forEach((vt) => {
            if (vt.GodownId !== null && vt.GodownId !== undefined && godownIdSet.has(Number(vt.GodownId))) {
                const name = (vt.label || vt.Voucher_Type || "").trim();
                if (name) set.add(name.toUpperCase());
            }
        });
        return set;
    }, [selectedGodownIds, voucherTypes]);

    // Dropdown options for ReportFilterDrawer
    const godownDropdownOptions = useMemo(() => {
        return godowns.map((g) => ({
            label: g.Godown_Name,
            value: Number(g.Godown_Id)
        }));
    }, [godowns]);

    // Normalize Category helper
    const normalizeCategory = (cat: string) => {
        if (!cat) return "";
        const c = cat.toUpperCase();
        if (c === "PROCESS") return "ADJUSTMENTS";
        return c;
    };

    // Process & Aggregate raw 2nd API items for a voucher
    // Deduplicates rows across multiple cost centers for the same invoice item, then applies columnFilters and groups by Brand + Item_Name_Modified + enabled columns
    const processVoucherItems = (rawItems: any[], enabledCols: ColumnConfig[]) => {
        if (!rawItems || rawItems.length === 0) return [];

        // 1. Deduplicate by unique invoice item: (Inv_No or Voucher_Ref_Id, Ref_Id, Stock_Item)
        const uniqueItemMap = new Map<string, any>();

        rawItems.forEach(row => {
            const invRef = row.Inv_No ? String(row.Inv_No).trim().toUpperCase() : String(row.Voucher_Ref_Id || "");
            const refId = String(row.Ref_Id || row.Stock_Item || "");
            const uniqueKey = `${invRef}___${refId}___${String(row.Stock_Item || "").trim().toUpperCase()}`;

            if (!uniqueItemMap.has(uniqueKey)) {
                uniqueItemMap.set(uniqueKey, { ...row });
            }
        });

        let deduplicatedItems = Array.from(uniqueItemMap.values());

        // 2. Filter by active columnFilters if any
        const activeFilterEntries = Object.entries(columnFilters).filter(
            ([, vals]) => Array.isArray(vals) && vals.length > 0
        );

        if (activeFilterEntries.length > 0) {
            deduplicatedItems = deduplicatedItems.filter(item => {
                return activeFilterEntries.every(([key, selectedVals]) => {
                    const itemVal = String(item[key] ?? "").trim();
                    return selectedVals!.includes(itemVal);
                });
            });
        }

        if (deduplicatedItems.length === 0) return [];

        // 3. Group items dynamically by active enabled columns
        const groupMap = new Map<string, any>();

        deduplicatedItems.forEach(item => {
            const compositeKey = enabledCols.length > 0
                ? enabledCols.map(c => `${c.key}:${String(item[c.key] ?? "-").trim()}`).join(":::")
                : "__ALL__";

            if (!groupMap.has(compositeKey)) {
                const entry: Record<string, any> = {
                    Bill_Qty: 0,
                    Act_Qty: 0,
                    itemCount: 0,
                    invoices: new Set<string>()
                };

                // Copy all available fields from first encounter
                Object.keys(item).forEach(k => {
                    entry[k] = item[k];
                });

                // Ensure all enabled column values are explicitly captured
                enabledCols.forEach(c => {
                    entry[c.key] = item[c.key] ?? "-";
                });

                groupMap.set(compositeKey, entry);
            }

            const target = groupMap.get(compositeKey);
            target.Bill_Qty += Number(item.Bill_Qty) || 0;
            target.Act_Qty += Number(item.Act_Qty ?? item.Bill_Qty) || 0;
            target.itemCount += 1;
            if (item.Inv_No) {
                target.invoices.add(String(item.Inv_No).trim());
            }
        });

        const result = Array.from(groupMap.values()).map(entry => ({
            ...entry,
            invoiceCount: entry.invoices.size,
            splitKgs: qtyType === "Act_Qty" ? entry.Act_Qty : entry.Bill_Qty
        }));

        // Sort dynamically by enabled columns in their configured order
        result.sort((a, b) => {
            for (const col of enabledCols) {
                const valA = String(a[col.key] ?? "");
                const valB = String(b[col.key] ?? "");
                const cmp = valA.localeCompare(valB, undefined, { numeric: true, sensitivity: "base" });
                if (cmp !== 0) return cmp;
            }
            return 0;
        });

        return result;
    };

    // Build hierarchical Categories and Groups from 1st API
    const tableCategories = useMemo(() => {
        const catList: ("INWARDS" | "ADJUSTMENTS" | "OUTWARDS")[] = ["INWARDS", "ADJUSTMENTS", "OUTWARDS"];
        const categoryGroupsMap: Record<string, any[]> = {
            INWARDS: [],
            ADJUSTMENTS: [],
            OUTWARDS: []
        };

        const activeFilterEntries = Object.entries(columnFilters).filter(
            ([, vals]) => Array.isArray(vals) && vals.length > 0
        );
        const hasActiveFilters = activeFilterEntries.length > 0;

        // 1. Process configured groups
        groups.forEach(group => {
            const matchedRows = reportData.filter(row =>
                row.Group_Name === group.name &&
                normalizeCategory(row.Overall_GroupName) === group.parentCategory &&
                row.Voucher_Type &&
                (!linkedVoucherNames || linkedVoucherNames.has(String(row.Voucher_Type).trim().toUpperCase()))
            );
            if (matchedRows.length === 0) return;

            const voucherTypesData = matchedRows.map(r => {
                const totalTonnage = qtyType === "Act_Qty"
                    ? (r.Total_Act_Tonnage ?? r.Act_Total_Tonnage ?? r.Total_Act_Qty ?? r.Total_Tonnage ?? 0)
                    : (r.Total_Tonnage || 0);

                const vtId = r.Voucher_Type_Id || r.VoucherId || r.Vocher_Type_Id || voucherTypes.find((v: any) =>
                    String(v.label).trim().toUpperCase() === String(r.Voucher_Type).trim().toUpperCase() ||
                    String(v.Voucher_Type).trim().toUpperCase() === String(r.Voucher_Type).trim().toUpperCase()
                )?.Value;

                const vKey = `${group.parentCategory}_${group.name}_${r.Voucher_Type}`;
                let effectiveKgs = totalTonnage;
                let hasMatches = true;

                if (hasActiveFilters) {
                    const rawItems = itemData[vKey];
                    if (rawItems) {
                        const processed = processVoucherItems(rawItems, enabledItemColumns);
                        effectiveKgs = processed.reduce((sum, it) => sum + (Number(it.splitKgs) || 0), 0);
                        if (processed.length === 0) {
                            hasMatches = false;
                        }
                    }
                }

                return {
                    name: r.Voucher_Type,
                    voucherTypeId: vtId !== undefined && vtId !== null ? Number(vtId) : undefined,
                    baseKgs: effectiveKgs,
                    invoiceCount: r.Invoice_Count || 0,
                    itemCount: r.Item_Count || 0,
                    hasMatches
                };
            }).filter(vt => !hasActiveFilters || vt.hasMatches);

            if (voucherTypesData.length === 0) return;

            const groupKgsSum = voucherTypesData.reduce((sum, vt) => sum + vt.baseKgs, 0);

            const groupData = {
                groupName: group.name,
                groupKgs: groupKgsSum,
                voucherTypes: voucherTypesData
            };

            if (categoryGroupsMap[group.parentCategory]) {
                categoryGroupsMap[group.parentCategory].push(groupData);
            }
        });

        // 2. Process unassigned groups
        catList.forEach(cat => {
            const definedGroupNames = groups.filter(g => g.parentCategory === cat).map(g => g.name);

            const unassignedRows = reportData.filter(row =>
                normalizeCategory(row.Overall_GroupName) === cat &&
                row.Voucher_Type &&
                (row.Group_Name === null || row.Group_Name === undefined || !definedGroupNames.includes(row.Group_Name)) &&
                (!linkedVoucherNames || linkedVoucherNames.has(String(row.Voucher_Type).trim().toUpperCase()))
            );

            if (unassignedRows.length > 0) {
                const voucherTypesData = unassignedRows.map(r => {
                    const totalTonnage = qtyType === "Act_Qty"
                        ? (r.Total_Act_Tonnage ?? r.Act_Total_Tonnage ?? r.Total_Act_Qty ?? r.Total_Tonnage ?? 0)
                        : (r.Total_Tonnage || 0);

                    const vtId = r.Voucher_Type_Id || r.VoucherId || r.Vocher_Type_Id || voucherTypes.find((v: any) =>
                        String(v.label).trim().toUpperCase() === String(r.Voucher_Type).trim().toUpperCase() ||
                        String(v.Voucher_Type).trim().toUpperCase() === String(r.Voucher_Type).trim().toUpperCase()
                    )?.Value;

                    const vKey = `${cat}_UNASSIGNED_${r.Voucher_Type}`;
                    let effectiveKgs = totalTonnage;
                    let hasMatches = true;

                    if (hasActiveFilters) {
                        const rawItems = itemData[vKey];
                        if (rawItems) {
                            const processed = processVoucherItems(rawItems, enabledItemColumns);
                            effectiveKgs = processed.reduce((sum, it) => sum + (Number(it.splitKgs) || 0), 0);
                            if (processed.length === 0) {
                                hasMatches = false;
                            }
                        }
                    }

                    return {
                        name: r.Voucher_Type,
                        voucherTypeId: vtId !== undefined && vtId !== null ? Number(vtId) : undefined,
                        baseKgs: effectiveKgs,
                        invoiceCount: r.Invoice_Count || 0,
                        itemCount: r.Item_Count || 0,
                        hasMatches
                    };
                }).filter(vt => !hasActiveFilters || vt.hasMatches);

                if (voucherTypesData.length > 0) {
                    const groupKgsSum = voucherTypesData.reduce((sum, vt) => sum + vt.baseKgs, 0);
                    const unassignedGroupData = {
                        groupName: "UNASSIGNED",
                        groupKgs: groupKgsSum,
                        voucherTypes: voucherTypesData
                    };
                    categoryGroupsMap[cat].push(unassignedGroupData);
                }
            }
        });

        return catList.map(cat => ({
            name: cat,
            groups: categoryGroupsMap[cat],
            categoryKgs: categoryGroupsMap[cat].reduce((sum, g) => sum + g.groupKgs, 0)
        })).filter(cat => cat.groups.length > 0);
    }, [groups, reportData, voucherTypes, linkedVoucherNames, qtyType, columnFilters, itemData, enabledItemColumns]);

    // Grand Totals Calculation
    const grandTotals = useMemo(() => {
        let totalKgs = 0;
        tableCategories.forEach(cat => {
            totalKgs += cat.categoryKgs;
        });
        return { totalKgs };
    }, [tableCategories]);

    // Toggle expand voucher type & fetch 2nd API data on-demand
    const handleToggleExpandVoucher = async (parentCategory: string, groupName: string, voucherName: string, voucherTypeId?: number) => {
        const vKey = `${parentCategory}_${groupName}_${voucherName}`;
        const isExpanded = expandedVouchers.includes(vKey);

        if (!isExpanded) {
            setExpandedVouchers(prev => [...prev, vKey]);

            if (!itemData[vKey]) {
                setLoadingItems(prev => ({ ...prev, [vKey]: true }));
                try {
                    const matchedVoucher = voucherTypes.find((v: any) =>
                        String(v.label).trim().toUpperCase() === String(voucherName).trim().toUpperCase() ||
                        String(v.Voucher_Type).trim().toUpperCase() === String(voucherName).trim().toUpperCase() ||
                        String(v.Value) === String(voucherName)
                    );
                    const resolvedVoucherTypeId = voucherTypeId !== undefined && voucherTypeId !== null
                        ? Number(voucherTypeId)
                        : (matchedVoucher ? Number(matchedVoucher.Value) : undefined);

                    const res = await employeeReportGroupService.getEmployeeInvoicesWithItems({
                        Fromdate: fromDate,
                        Todate: toDate,
                        Overall_GroupName: parentCategory === "ADJUSTMENTS" ? "PROCESS" : parentCategory,
                        Group_Name: groupName,
                        Voucher_Type: resolvedVoucherTypeId !== undefined ? resolvedVoucherTypeId : voucherName
                    });

                    if (res.data?.success && res.data?.data) {
                        const rawItems = res.data.data;
                        setItemData(prev => ({
                            ...prev,
                            [vKey]: rawItems
                        }));
                        // Discover new dynamic columns from API response
                        setItemColumns(prev => discoverItemColumns(rawItems, prev));
                    } else {
                        setItemData(prev => ({ ...prev, [vKey]: [] }));
                    }
                } catch (err) {
                    console.error("Error loading voucher items:", err);
                    toast.error("Failed to load item split-up");
                } finally {
                    setLoadingItems(prev => ({ ...prev, [vKey]: false }));
                }
            }
        } else {
            setExpandedVouchers(prev => prev.filter(k => k !== vKey));
        }
    };

    // Fetch all missing 2nd API item data in batches of 4
    const fetchAllItemData = async () => {
        const allVKeys: { vKey: string; parentCategory: string; groupName: string; voucherName: string; voucherTypeId?: number }[] = [];
        tableCategories.forEach(cat => {
            cat.groups.forEach((g: any) => {
                g.voucherTypes.forEach((vt: any) => {
                    const vKey = `${cat.name}_${g.groupName}_${vt.name}`;
                    allVKeys.push({
                        vKey,
                        parentCategory: cat.name,
                        groupName: g.groupName,
                        voucherName: vt.name,
                        voucherTypeId: vt.voucherTypeId
                    });
                });
            });
        });

        const missing = allVKeys.filter(k => !itemData[k.vKey]);
        if (missing.length === 0) return allVKeys;

        const batchSize = 4;
        for (let i = 0; i < missing.length; i += batchSize) {
            const batch = missing.slice(i, i + batchSize);
            await Promise.all(
                batch.map(async (item) => {
                    setLoadingItems(prev => ({ ...prev, [item.vKey]: true }));
                    try {
                        const res = await employeeReportGroupService.getEmployeeInvoicesWithItems({
                            Fromdate: fromDate,
                            Todate: toDate,
                            Overall_GroupName: item.parentCategory === "ADJUSTMENTS" ? "PROCESS" : item.parentCategory,
                            Group_Name: item.groupName,
                            Voucher_Type: item.voucherTypeId !== undefined ? item.voucherTypeId : item.voucherName
                        });
                        if (res.data?.success && res.data?.data) {
                            setItemData(prev => ({ ...prev, [item.vKey]: res.data.data }));
                            setItemColumns(prev => discoverItemColumns(res.data.data, prev));
                        }
                    } catch (err) {
                        console.error(`Error loading items for ${item.vKey}:`, err);
                    } finally {
                        setLoadingItems(prev => ({ ...prev, [item.vKey]: false }));
                    }
                })
            );
        }
        return allVKeys;
    };

    // Auto-prefetch item data in background when report data arrives
    useEffect(() => {
        if (reportData.length > 0 && tableCategories.length > 0) {
            fetchAllItemData();
        }
    }, [reportData, tableCategories.length, fromDate, toDate]);

    // Header click handler for 2nd API columns
    const handleHeaderClick = (e: React.MouseEvent<HTMLElement>, key: string) => {
        setActiveHeader(key);
        setFilterAnchor(e.currentTarget);
        fetchAllItemData();
    };

    // Filter options for the active column header
    const filterOptions = useMemo(() => {
        if (!activeHeader) return [];
        const allItems = Object.values(itemData).flat();
        const vals = Array.from(
            new Set(
                allItems
                    .map(item => String(item[activeHeader] ?? "").trim())
                    .filter(Boolean)
            )
        );
        return vals.sort((a, b) => a.localeCompare(b));
    }, [activeHeader, itemData]);

    // When column filters are applied, auto-expand matching vouchers so items are immediately visible
    useEffect(() => {
        const activeFilters = Object.entries(columnFilters).filter(([, v]) => Array.isArray(v) && v.length > 0);
        if (activeFilters.length > 0) {
            const matchingKeys: string[] = [];
            tableCategories.forEach(cat => {
                cat.groups.forEach((g: any) => {
                    g.voucherTypes.forEach((vt: any) => {
                        const vKey = `${cat.name}_${g.groupName}_${vt.name}`;
                        matchingKeys.push(vKey);
                    });
                });
            });
            if (matchingKeys.length > 0) {
                setExpandedVouchers(matchingKeys);
            }
        }
    }, [columnFilters, tableCategories]);


    // Template Save Handler
    const handleQuickSave = async () => {
        try {
            if (!reportName.trim()) {
                toast.error("Enter Report Name");
                return;
            }

            const payloadColumns = itemColumns.map((c) => ({
                key: c.key,
                label: c.label,
                enabled: c.enabled,
                order: c.order,
                groupBy: 0,
                dataType: "string"
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
                    abstractSP: "SP_Get_EmployeeReport_EmployeeInvoices_with_ITEM",
                    expandedSP: "SP_Get_EmployeeReport_EmployeeInvoices_with_ITEM",
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

    // Template Load Handler
    const handleLoadTemplate = async (templateId: number) => {
        try {
            setSelectedTemplateId(templateId);
            setExpandedVouchers([]);
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

            const parsedTemplateCols: ColumnConfig[] = templateCols
                .filter((matched: any) => {
                    const key = matched.key || matched.Key || matched.Column_Name || matched.ColumnName || "";
                    return key && !EXCLUDED_ITEM_KEYS.has(key);
                })
                .map((matched: any, idx: number) => {
                    const key = matched.key || matched.Key || matched.Column_Name || matched.ColumnName || "";
                    const label = matched.label || matched.Label || matched.Alias_Name || key;
                    const enabled = matched.enabled ?? matched.Enabled ?? false;
                    const order = matched.order ?? matched.Order ?? idx;

                    return {
                        key,
                        label,
                        enabled,
                        order
                    };
                });

            // Merge with existing discovered columns
            const updatedCols = [...parsedTemplateCols];
            itemColumns.forEach(col => {
                if (EXCLUDED_ITEM_KEYS.has(col.key)) return;
                const exists = updatedCols.some(c => c.key.toUpperCase() === col.key.toUpperCase());
                if (!exists) {
                    updatedCols.push({
                        ...col,
                        enabled: false,
                        order: updatedCols.length
                    });
                }
            });

            updatedCols.sort((a, b) => a.order - b.order);
            setItemColumns(updatedCols.map((c, index) => ({ ...c, order: index })));
            if (data.filters?.columnFilters) {
                setColumnFilters(data.filters.columnFilters);
            }

            toast.success("Template Loaded Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to load template ❌");
        }
    };

    // Reset Template
    const handleClearTemplate = () => {
        setSelectedTemplateId(null);
        setReportName("");
        setSelectedGodownIds([]);
        setTempSelectedGodownIds([]);
        setColumnFilters({});
        setItemColumns(DEFAULT_ITEM_COLUMNS);
        setExpandedVouchers([]);
    };

    // Excel Export: Summary
    const handleSummaryExportExcel = () => {
        const rows: any[] = [];
        tableCategories.forEach(cat => {
            cat.groups.forEach((g: any) => {
                g.voucherTypes.forEach((vt: any) => {
                    rows.push({
                        Category: cat.name,
                        Group: g.groupName,
                        "Group Kgs": g.groupKgs,
                        "Voucher Type": vt.name,
                        "Voucher Kgs": vt.baseKgs
                    });
                });
            });
        });

        // Add Grand Total
        rows.push({
            Category: "Total",
            Group: "",
            "Group Kgs": "",
            "Voucher Type": "Grand Total",
            "Voucher Kgs": grandTotals.totalKgs
        });

        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");
        XLSX.writeFile(workbook, `Group_Wise_Item_Summary_${fromDate}_to_${toDate}.xlsx`);
    };

    // Excel Export: Detailed with Item Split-up
    const handleDetailedExportExcel = async () => {
        toast.info("Preparing detailed item export...");
        // Ensure all item data is fetched without modifying UI expanded rows
        await fetchAllItemData();

        const rows: any[] = [];
        tableCategories.forEach(cat => {
            cat.groups.forEach((g: any) => {
                g.voucherTypes.forEach((vt: any) => {
                    const vKey = `${cat.name}_${g.groupName}_${vt.name}`;
                    const rawItems = itemData[vKey] || [];
                    const processed = processVoucherItems(rawItems, enabledItemColumns);

                    if (processed.length > 0) {
                        processed.forEach(item => {
                            const rowObj: Record<string, any> = {
                                Category: cat.name,
                                Group: g.groupName,
                                "Voucher Type": vt.name
                            };
                            enabledItemColumns.forEach(col => {
                                rowObj[col.label] = item[col.key] ?? "-";
                            });
                            rowObj["Split Kgs"] = item.splitKgs;
                            rows.push(rowObj);
                        });
                    } else {
                        const rowObj: Record<string, any> = {
                            Category: cat.name,
                            Group: g.groupName,
                            "Voucher Type": vt.name
                        };
                        enabledItemColumns.forEach(col => {
                            rowObj[col.label] = "-";
                        });
                        rowObj["Split Kgs"] = vt.baseKgs;
                        rows.push(rowObj);
                    }
                });
            });
        });

        // Total
        const totalObj: Record<string, any> = {
            Category: "Total",
            Group: "",
            "Voucher Type": "Grand Total"
        };
        enabledItemColumns.forEach(col => {
            totalObj[col.label] = "";
        });
        totalObj["Split Kgs"] = grandTotals.totalKgs;
        rows.push(totalObj);

        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Item Split-up");
        XLSX.writeFile(workbook, `Group_Wise_Item_Detailed_${fromDate}_to_${toDate}.xlsx`);
    };

    // PDF Export: Summary
    const handleSummaryExportPDF = () => {
        const doc = new jsPDF("l", "pt", "a4");
        doc.setFontSize(14);
        doc.text(`Group Wise Item Report (${fromDate} to ${toDate}) - Summary`, 40, 30);

        const tableBody: any[] = [];
        tableCategories.forEach(cat => {
            cat.groups.forEach((g: any) => {
                g.voucherTypes.forEach((vt: any) => {
                    tableBody.push([
                        cat.name,
                        g.groupName,
                        g.groupKgs.toLocaleString(),
                        vt.name,
                        vt.baseKgs.toLocaleString()
                    ]);
                });
            });
        });

        tableBody.push([
            "Total",
            "",
            "",
            "Grand Total",
            grandTotals.totalKgs.toLocaleString()
        ]);

        autoTable(doc, {
            head: [["Category", "Group", "Group Kgs", "Voucher Type", "Voucher Kgs"]],
            body: tableBody,
            startY: 45,
            theme: "grid",
            styles: { fontSize: 8 },
            headStyles: { fillColor: [30, 58, 138], textColor: 255 }
        });

        doc.save(`Group_Wise_Item_Summary_${fromDate}_to_${toDate}.pdf`);
    };

    // PDF Export: Detailed
    const handleDetailedExportPDF = async () => {
        toast.info("Preparing detailed item PDF...");
        // Ensure all item data is fetched without modifying UI expanded rows
        await fetchAllItemData();

        const doc = new jsPDF("l", "pt", "a4");
        doc.setFontSize(14);
        doc.text(`Group Wise Item Report (${fromDate} to ${toDate}) - Detailed Items`, 40, 30);

        const headCols = ["Category", "Group", "Voucher Type", ...enabledItemColumns.map(c => c.label), "Split Kgs"];
        const tableBody: any[] = [];

        tableCategories.forEach(cat => {
            cat.groups.forEach((g: any) => {
                g.voucherTypes.forEach((vt: any) => {
                    const vKey = `${cat.name}_${g.groupName}_${vt.name}`;
                    const rawItems = itemData[vKey] || [];
                    const processed = processVoucherItems(rawItems, enabledItemColumns);

                    if (processed.length > 0) {
                        processed.forEach(item => {
                            const row = [
                                cat.name,
                                g.groupName,
                                vt.name,
                                ...enabledItemColumns.map(col => String(item[col.key] ?? "-")),
                                item.splitKgs.toLocaleString()
                            ];
                            tableBody.push(row);
                        });
                    } else {
                        const row = [
                            cat.name,
                            g.groupName,
                            vt.name,
                            ...enabledItemColumns.map(() => "-"),
                            vt.baseKgs.toLocaleString()
                        ];
                        tableBody.push(row);
                    }
                });
            });
        });

        tableBody.push([
            "Total",
            "",
            "Grand Total",
            ...enabledItemColumns.map(() => ""),
            grandTotals.totalKgs.toLocaleString()
        ]);

        autoTable(doc, {
            head: [headCols],
            body: tableBody,
            startY: 45,
            theme: "grid",
            styles: { fontSize: 7 },
            headStyles: { fillColor: [30, 58, 138], textColor: 255 }
        });

        doc.save(`Group_Wise_Item_Detailed_${fromDate}_to_${toDate}.pdf`);
    };

    // Styling constants
    const headerCellHeight = 38;
    const headerCellSx = {
        color: "#ffffff",
        fontWeight: 700,
        height: headerCellHeight,
        maxHeight: headerCellHeight,
        minHeight: headerCellHeight,
        lineHeight: "38px",
        py: 0,
        borderRight: "1px solid #2448b2",
        whiteSpace: "nowrap" as const,
        fontSize: "0.8rem",
        bgcolor: "#1E3A8A",
        position: "sticky" as const,
        top: 0,
        zIndex: 4
    };

    const totalCellSx = {
        fontWeight: 800,
        color: "#1e3a8a",
        fontSize: "0.825rem",
        py: 0.8,
        bgcolor: "#bfdbfe",
        borderRight: "1px solid #93c5fd",
        borderBottom: "none",
        whiteSpace: "nowrap" as const,
        position: "sticky" as const,
        top: headerCellHeight,
        zIndex: 3
    };

    return (
        <Box sx={{ display: "flex", flexDirection: "column", height: "100%", width: "100%", overflow: "hidden" }}>
            <PageHeader
                parentReportName={parentReportName}
                onExportExcel={() => {
                    setExportType("excel");
                    setExportModalOpen(true);
                }}
                onExportPDF={() => {
                    setExportType("pdf");
                    setExportModalOpen(true);
                }}
                onReportChange={(template) => {
                    if (!template || !template.Report_Id) {
                        handleClearTemplate();
                    } else {
                        setSelectedTemplateId(template.Report_Id);
                        setReportName(template.Report_Name || "");
                        handleLoadTemplate(template.Report_Id);
                    }
                }}
                onQuickSave={() => {
                    if (!selectedTemplateId) {
                        setReportName("");
                    }
                    setSaveDialogOpen(true);
                }}
                settingsSlot={
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
                dropdownLabel="Godown"
                dropdownMultiple={true}
                dropdownValue={tempSelectedGodownIds}
                dropdownOptions={godownDropdownOptions}
                onDropdownChange={(val) => {
                    const ids = Array.isArray(val) ? val.map(Number) : [];
                    setTempSelectedGodownIds(ids);
                }}
                onApply={() => {
                    setFromDate(tempFromDate);
                    setToDate(tempToDate);
                    setSelectedGodownIds(tempSelectedGodownIds);
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

            <Box px={2} pb={1} pt={1} sx={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", overflow: "hidden" }}>
                {/* Active Godown & Column Filter Chips */}
                {(selectedGodownIds.length > 0 || Object.entries(columnFilters).some(([, v]) => Array.isArray(v) && v.length > 0)) && (
                    <Box display="flex" alignItems="center" flexWrap="wrap" gap={1} mb={1}>
                        {selectedGodownIds.length > 0 && (
                            <>
                                <Typography variant="body2" sx={{ fontWeight: 600, color: "#475569", fontSize: "0.75rem" }}>
                                    Godowns:
                                </Typography>
                                {selectedGodownIds.map((gid) => {
                                    const matched = godowns.find((g) => Number(g.Godown_Id) === gid);
                                    return (
                                        <Chip
                                            key={gid}
                                            label={matched?.Godown_Name || gid}
                                            size="small"
                                            onDelete={() => {
                                                setSelectedGodownIds(prev => prev.filter(id => id !== gid));
                                                setTempSelectedGodownIds(prev => prev.filter(id => id !== gid));
                                            }}
                                            sx={{
                                                bgcolor: "#e2e8f0",
                                                fontWeight: 600,
                                                fontSize: "0.75rem",
                                                color: "#334155"
                                            }}
                                        />
                                    );
                                })}
                            </>
                        )}

                        {Object.entries(columnFilters).map(([key, vals]) => {
                            if (!Array.isArray(vals) || vals.length === 0) return null;
                            const colLabel = itemColumns.find(c => c.key === key)?.label || key;
                            const chipText = vals.length <= 2
                                ? `${colLabel}: ${vals.join(", ")}`
                                : `${colLabel}: ${vals.length} selected`;
                            return (
                                <Chip
                                    key={key}
                                    label={chipText}
                                    size="small"
                                    onDelete={() => {
                                        setColumnFilters(prev => {
                                            const copy = { ...prev };
                                            delete copy[key];
                                            return copy;
                                        });
                                    }}
                                    sx={{
                                        bgcolor: "#e0f2fe",
                                        fontWeight: 600,
                                        fontSize: "0.75rem",
                                        color: "#0369a1"
                                    }}
                                />
                            );
                        })}

                        <Chip
                            label="Clear All Filters"
                            size="small"
                            onClick={() => {
                                setSelectedGodownIds([]);
                                setTempSelectedGodownIds([]);
                                setColumnFilters({});
                            }}
                            sx={{
                                bgcolor: "#fee2e2",
                                fontWeight: 600,
                                fontSize: "0.75rem",
                                color: "#ef4444",
                                cursor: "pointer",
                                "&:hover": { bgcolor: "#fecaca" }
                            }}
                        />
                    </Box>
                )}

                {/* Main Table */}
                {loading ? (
                    <Box display="flex" flexDirection="column" alignItems="center" justifyContent="center" height="60vh" gap={2}>
                        <CircularProgress size={40} />
                        <Typography variant="body1" color="text.secondary" fontWeight={500}>
                            Loading report data...
                        </Typography>
                    </Box>
                ) : (
                    <TableContainer component={Paper} elevation={1} sx={{ flex: 1, minHeight: 0, borderRadius: 2, border: "1px solid #cbd5e1", overflow: "auto" }}>
                        <Table size="small" stickyHeader>
                            <TableHead>
                                <TableRow sx={{ bgcolor: "#1E3A8A" }}>
                                    <TableCell sx={headerCellSx} align="center">
                                        Category
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Groups
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Group Kgs
                                    </TableCell>
                                    <TableCell sx={headerCellSx}>
                                        Voucher Type
                                    </TableCell>
                                    <TableCell sx={headerCellSx} align="center">
                                        Voucher Kgs
                                    </TableCell>
                                    {enabledItemColumns.map(col => {
                                        return (
                                            <TableCell
                                                key={col.key}
                                                sx={{
                                                    ...headerCellSx,
                                                    cursor: "pointer",
                                                    userSelect: "none",
                                                    "&:hover": { bgcolor: "#2448b2" }
                                                }}
                                                align={col.key === "Brokerage" || col.key === "Coolie" ? "center" : "left"}
                                                onClick={(e) => handleHeaderClick(e, col.key)}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="space-between" gap={0.5}>
                                                    <span>{col.label}</span>
                                                </Box>
                                            </TableCell>
                                        );
                                    })}
                                    <TableCell sx={{ ...headerCellSx, borderRight: "none" }} align="center">
                                        Split Kgs
                                    </TableCell>
                                </TableRow>

                                {/* Grand Total Row */}
                                <TableRow sx={{ bgcolor: "#bfdbfe" }}>
                                    <TableCell colSpan={2} sx={totalCellSx} align="center">
                                        Total
                                    </TableCell>
                                    <TableCell sx={totalCellSx} align="center">
                                        {grandTotals.totalKgs.toLocaleString()}
                                    </TableCell>
                                    <TableCell sx={totalCellSx}>
                                        Grand Total
                                    </TableCell>
                                    <TableCell sx={totalCellSx} align="center">
                                        {grandTotals.totalKgs.toLocaleString()}
                                    </TableCell>
                                    {enabledItemColumns.map(col => (
                                        <TableCell key={col.key} sx={totalCellSx} />
                                    ))}
                                    <TableCell sx={{ ...totalCellSx, borderRight: "none" }} align="center">
                                        {grandTotals.totalKgs.toLocaleString()}
                                    </TableCell>
                                </TableRow>
                            </TableHead>

                            <TableBody>
                                {tableCategories.length > 0 ? (
                                    tableCategories.map((category) => {
                                        // Calculate total spans for category including expanded items
                                        let categorySpan = 0;
                                        category.groups.forEach((group: any) => {
                                            group.voucherTypes.forEach((vt: any) => {
                                                categorySpan += 1;
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                if (expandedVouchers.includes(vKey)) {
                                                    const rawItems = itemData[vKey] || [];
                                                    const processed = processVoucherItems(rawItems, enabledItemColumns);
                                                    categorySpan += (processed.length > 0 ? processed.length : (loadingItems[vKey] ? 1 : 1));
                                                }
                                            });
                                        });

                                        let isFirstCategoryRow = true;

                                        return category.groups.map((group: any) => {
                                            let groupSpan = 0;
                                            group.voucherTypes.forEach((vt: any) => {
                                                groupSpan += 1;
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                if (expandedVouchers.includes(vKey)) {
                                                    const rawItems = itemData[vKey] || [];
                                                    const processed = processVoucherItems(rawItems, enabledItemColumns);
                                                    groupSpan += (processed.length > 0 ? processed.length : (loadingItems[vKey] ? 1 : 1));
                                                }
                                            });

                                            let isFirstGroupRow = true;

                                            return group.voucherTypes.map((vt: any) => {
                                                const vKey = `${category.name}_${group.groupName}_${vt.name}`;
                                                const isExpanded = expandedVouchers.includes(vKey);
                                                const isLoadingItemData = loadingItems[vKey];
                                                const rawItems = itemData[vKey] || [];
                                                const processedItems = processVoucherItems(rawItems, enabledItemColumns);

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
                                                                <Typography variant="body2" fontWeight={800} mb={0.5}>
                                                                    {category.name}
                                                                </Typography>
                                                                <Typography variant="caption" fontWeight={800} color="#0369a1">
                                                                    {category.categoryKgs.toLocaleString()}
                                                                </Typography>
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
                                                                    fontSize: "0.825rem"
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
                                                                    fontSize: "0.825rem"
                                                                }}
                                                                align="center"
                                                            >
                                                                {group.groupKgs.toLocaleString()}
                                                            </TableCell>
                                                        )}

                                                        {/* Voucher Type Column with expansion click */}
                                                        <TableCell
                                                            sx={{
                                                                fontWeight: 700,
                                                                borderRight: "1px solid #cbd5e1",
                                                                color: "#334155",
                                                                cursor: "pointer",
                                                                "&:hover": { color: "#1e3a8a", bgcolor: "#f1f5f9" },
                                                                fontSize: "0.825rem",
                                                                py: 0.8
                                                            }}
                                                            onClick={() => handleToggleExpandVoucher(category.name, group.groupName, vt.name, vt.voucherTypeId)}
                                                        >
                                                            <Box display="flex" alignItems="center" gap={0.5}>
                                                                {isExpanded ? <KeyboardArrowUpIcon fontSize="small" /> : <KeyboardArrowDownIcon fontSize="small" />}
                                                                <span>{vt.name}</span>
                                                            </Box>
                                                        </TableCell>

                                                        {/* Voucher Kgs */}
                                                        <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #cbd5e1", fontSize: "0.825rem", py: 0.8 }} align="center">
                                                            {vt.baseKgs.toLocaleString()}
                                                        </TableCell>

                                                        {/* Placeholder cells for item columns in the main voucher summary row */}
                                                        {enabledItemColumns.map(col => (
                                                            <TableCell key={col.key} sx={{ color: "#94a3b8", fontStyle: "italic", borderRight: "1px solid #cbd5e1", py: 0.8 }} align="center">
                                                                -
                                                            </TableCell>
                                                        ))}

                                                        <TableCell sx={{ fontWeight: 700, color: "#64748b", borderRight: "none", py: 0.8 }} align="center">
                                                            {vt.baseKgs.toLocaleString()}
                                                        </TableCell>
                                                    </TableRow>
                                                );

                                                isFirstCategoryRow = false;
                                                isFirstGroupRow = false;

                                                // Expanded Item Rows
                                                const itemRows = isExpanded ? (
                                                    isLoadingItemData ? (
                                                        <TableRow key={`${vKey}_loading`}>
                                                            <TableCell colSpan={enabledItemColumns.length + 3} sx={{ py: 1.5, textAlign: "center", bgcolor: "#fafafa" }}>
                                                                <Box display="flex" alignItems="center" justifyContent="center" gap={1}>
                                                                    <CircularProgress size={18} />
                                                                    <Typography variant="body2" color="text.secondary" fontSize="0.775rem">
                                                                        Loading items for {vt.name}...
                                                                    </Typography>
                                                                </Box>
                                                            </TableCell>
                                                        </TableRow>
                                                    ) : processedItems.length > 0 ? (
                                                        processedItems.map((item: any, itemIdx: number) => (
                                                            <TableRow
                                                                key={`${vKey}_item_${itemIdx}`}
                                                                sx={{
                                                                    bgcolor: itemIdx % 2 === 0 ? "#ffffff" : "#fdfdfe",
                                                                    "&:hover": { bgcolor: "#f8fafc" },
                                                                    borderBottom: "1px solid #e2e8f0"
                                                                }}
                                                            >
                                                                <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.5 }} />
                                                                <TableCell sx={{ borderRight: "1px solid #cbd5e1", py: 0.5 }} />

                                                                {/* Render all enabled item columns */}
                                                                {enabledItemColumns.map((col) => {
                                                                    const val = item[col.key];
                                                                    const isNumeric = col.key === "Brokerage" || col.key === "Coolie";
                                                                    return (
                                                                        <TableCell
                                                                            key={col.key}
                                                                            sx={{
                                                                                color: col.key === "Item_Name_Modified" ? "#1e3a8a" : "#334155",
                                                                                fontWeight: col.key === "Item_Name_Modified" || col.key === "Brand" || col.key === "Stock_Group" ? 600 : 400,
                                                                                borderRight: "1px solid #cbd5e1",
                                                                                py: 0.5,
                                                                                fontSize: "0.775rem",
                                                                                pl: col.key === "Brand" ? 3 : (isNumeric ? 1 : 2)
                                                                            }}
                                                                            align={isNumeric ? "center" : "left"}
                                                                        >
                                                                            {val !== undefined && val !== null && val !== "" ? String(val) : "-"}
                                                                        </TableCell>
                                                                    );
                                                                })}

                                                                {/* Split Kgs cell */}
                                                                <TableCell
                                                                    sx={{
                                                                        fontWeight: 700,
                                                                        color: "#0f766e",
                                                                        bgcolor: "#f0fdfa",
                                                                        borderRight: "none",
                                                                        py: 0.5,
                                                                        fontSize: "0.8rem"
                                                                    }}
                                                                    align="center"
                                                                >
                                                                    {item.splitKgs.toLocaleString()}
                                                                </TableCell>
                                                            </TableRow>
                                                        ))
                                                    ) : (
                                                        <TableRow key={`${vKey}_no_data`}>
                                                            <TableCell colSpan={enabledItemColumns.length + 3} sx={{ py: 1, textAlign: "center", bgcolor: "#fbfbfe", color: "text.secondary", fontSize: "0.75rem", fontStyle: "italic" }}>
                                                                No individual item records found for {vt.name}
                                                            </TableCell>
                                                        </TableRow>
                                                    )
                                                ) : null;

                                                return (
                                                    <React.Fragment key={vt.name}>
                                                        {mainRow}
                                                        {itemRows}
                                                    </React.Fragment>
                                                );
                                            });
                                        });
                                    })
                                ) : (
                                    <TableRow>
                                        <TableCell colSpan={enabledItemColumns.length + 6} align="center" sx={{ py: 4 }}>
                                            <Typography variant="body2" color="text.secondary">
                                                No group data found for the selected dates and filters.
                                            </Typography>
                                        </TableCell>
                                    </TableRow>
                                )}
                            </TableBody>
                        </Table>
                    </TableContainer>
                )}
            </Box>

            {/* ******* COLUMN SETTINGS MENU ******* */}
            <Menu
                anchorEl={settingsAnchor}
                open={Boolean(settingsAnchor)}
                onClose={() => setSettingsAnchor(null)}
                PaperProps={{
                    sx: {
                        width: 320,
                        maxHeight: 480,
                        px: 2,
                        py: 1.5,
                        boxShadow: "0px 4px 20px rgba(0, 0, 0, 0.1)",
                        borderRadius: 2
                    }
                }}
            >
                <Box sx={{ maxHeight: 420, overflowY: "auto" }}>
                    <Box display="flex" justifyContent="space-between" alignItems="center" mb={1} pb={0.5} borderBottom="1px solid #e2e8f0">
                        <Typography variant="subtitle2" fontWeight={700} color="#1E3A8A">
                            Column Settings
                        </Typography>
                        <Button
                            size="small"
                            variant="text"
                            onClick={() => {
                                setItemColumns(DEFAULT_ITEM_COLUMNS);
                                setColumnFilters({});
                            }}
                            sx={{ textTransform: "none", fontWeight: 600, minWidth: 0, p: 0, color: "#1E3A8A" }}
                        >
                            Reset
                        </Button>
                    </Box>

                    <Typography variant="caption" fontWeight="bold" color="textSecondary" display="block" mb={1}>
                        Enabled Columns
                    </Typography>

                    <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
                        <SortableContext
                            items={itemColumns.filter(c => c.enabled).sort((a, b) => a.order - b.order).map(c => c.key)}
                            strategy={verticalListSortingStrategy}
                        >
                            {itemColumns
                                .filter(c => c.enabled)
                                .sort((a, b) => a.order - b.order)
                                .map(col => (
                                    <SortableColumnRow
                                        key={col.key}
                                        column={col}
                                        onToggle={handleToggleColumn}
                                        hasActiveFilter={columnFilters[col.key] !== undefined && (columnFilters[col.key]?.length ?? 0) > 0}
                                    />
                                ))}
                        </SortableContext>
                    </DndContext>

                    <Box mt={2}>
                        <Typography variant="caption" fontWeight="bold" color="textSecondary" display="block" mb={1}>
                            Disabled Columns
                        </Typography>

                        {itemColumns
                            .filter(c => !c.enabled)
                            .map(col => (
                                <Box
                                    key={col.key}
                                    display="flex"
                                    justifyContent="space-between"
                                    alignItems="center"
                                    py={0.5}
                                    pl={1}
                                >
                                    <Box display="flex" alignItems="center" gap={1}>
                                        <Typography fontSize="0.75rem">{col.label}</Typography>
                                        {columnFilters[col.key] !== undefined && (columnFilters[col.key]?.length ?? 0) > 0 && (
                                            <Tooltip title="Header filter enabled">
                                                <FilterAltIcon fontSize="small" color="primary" sx={{ fontSize: "0.85rem" }} />
                                            </Tooltip>
                                        )}
                                    </Box>

                                    <Switch
                                        size="medium"
                                        checked={col.enabled}
                                        onChange={() => handleToggleColumn(col.key)}
                                        sx={{
                                            "& .MuiSwitch-switchBase.Mui-checked": {
                                                color: "#1E3A8A",
                                            },
                                            "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track": {
                                                backgroundColor: "#b5b9c4",
                                            },
                                            "& .MuiSwitch-track": {
                                                backgroundColor: "#CBD5E1",
                                            },
                                        }}
                                    />
                                </Box>
                            ))}
                    </Box>
                </Box>
            </Menu>

            {/* ******* TEMPLATE SAVE DIALOG ******* */}
            <Dialog open={saveDialogOpen} onClose={() => setSaveDialogOpen(false)} maxWidth="xs" fullWidth>
                <DialogTitle sx={{ color: "#1E3A8A", fontWeight: "bold" }}>
                    {selectedTemplateId ? "Update Template" : "Save as Template"}
                </DialogTitle>
                <DialogContent>
                    <TextField
                        autoFocus
                        margin="dense"
                        label="Template Name"
                        fullWidth
                        size="small"
                        value={reportName}
                        onChange={(e) => setReportName(e.target.value)}
                        sx={{ mt: 1 }}
                    />
                </DialogContent>
                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button onClick={() => setSaveDialogOpen(false)} color="warning" sx={{ textTransform: "none" }}>
                        Cancel
                    </Button>
                    <Button
                        onClick={handleQuickSave}
                        variant="contained"
                        sx={{
                            bgcolor: "#1E3A8A",
                            textTransform: "none",
                            fontWeight: 600,
                            "&:hover": { bgcolor: "#152a66" }
                        }}
                    >
                        Save
                    </Button>
                </DialogActions>
            </Dialog>



            {/* Export Dialog Modal */}
            <Dialog
                open={exportModalOpen}
                onClose={() => setExportModalOpen(false)}
                PaperProps={{
                    sx: {
                        borderRadius: "16px",
                        padding: "8px",
                        boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)",
                        width: "100%",
                        maxWidth: "440px"
                    }
                }}
            >
                <DialogTitle sx={{ m: 0, p: 2, display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: "1px solid #f1f5f9" }}>
                    <Typography variant="h6" fontWeight="bold" color="#1e3a8a">
                        Export Report
                    </Typography>
                    <IconButton onClick={() => setExportModalOpen(false)} size="small">
                        <CloseIcon />
                    </IconButton>
                </DialogTitle>

                <DialogContent sx={{ p: 3, pt: 3 }}>
                    <Typography variant="body2" color="text.secondary" mb={3}>
                        Select how you would like to export your Group Wise Item report.
                    </Typography>

                    <Box display="flex" flexDirection="column" gap={2}>
                        {/* Option 1: Summary */}
                        <Box
                            onClick={() => setExportDetails(false)}
                            sx={{
                                border: `2px solid ${!exportDetails ? "#3b82f6" : "#e2e8f0"}`,
                                borderRadius: "12px",
                                p: 2,
                                cursor: "pointer",
                                bgcolor: !exportDetails ? "#eff6ff" : "transparent",
                                transition: "all 0.2s",
                                "&:hover": {
                                    borderColor: "#3b82f6",
                                    bgcolor: "#f8fafc"
                                }
                            }}
                        >
                            <Typography variant="subtitle1" fontWeight="bold" color={!exportDetails ? "#1e40af" : "#1e293b"}>
                                Summary Report
                            </Typography>
                            <Typography variant="body2" color="text.secondary" mt={0.5}>
                                Exports Category, Group, and Voucher Type overview summary with totals.
                            </Typography>
                        </Box>

                        {/* Option 2: Detailed */}
                        <Box
                            onClick={() => setExportDetails(true)}
                            sx={{
                                border: `2px solid ${exportDetails ? "#3b82f6" : "#e2e8f0"}`,
                                borderRadius: "12px",
                                p: 2,
                                cursor: "pointer",
                                bgcolor: exportDetails ? "#eff6ff" : "transparent",
                                transition: "all 0.2s",
                                "&:hover": {
                                    borderColor: "#3b82f6",
                                    bgcolor: "#f8fafc"
                                }
                            }}
                        >
                            <Typography variant="subtitle1" fontWeight="bold" color={exportDetails ? "#1e40af" : "#1e293b"}>
                                Detailed Report
                            </Typography>
                            <Typography variant="body2" color="text.secondary" mt={0.5}>
                                Includes Brand, Item Name, and individual item split-up Kgs breakdown.
                            </Typography>
                        </Box>
                    </Box>

                    <Box display="flex" justifyContent="flex-end" gap={1.5} mt={4}>
                        <Button
                            onClick={() => setExportModalOpen(false)}
                            sx={{
                                color: "#64748b",
                                textTransform: "none",
                                fontWeight: "semibold",
                                borderRadius: "8px"
                            }}
                        >
                            Cancel
                        </Button>
                        <Button
                            variant="contained"
                            onClick={() => {
                                setExportModalOpen(false);
                                if (exportType === "excel") {
                                    if (exportDetails) {
                                        handleDetailedExportExcel();
                                    } else {
                                        handleSummaryExportExcel();
                                    }
                                } else if (exportType === "pdf") {
                                    if (exportDetails) {
                                        handleDetailedExportPDF();
                                    } else {
                                        handleSummaryExportPDF();
                                    }
                                }
                            }}
                            sx={{
                                bgcolor: "#1e3a8a",
                                textTransform: "none",
                                fontWeight: "semibold",
                                borderRadius: "8px",
                                px: 3,
                                "&:hover": { bgcolor: "#152a66" }
                            }}
                        >
                            Export
                        </Button>
                    </Box>
                </DialogContent>
            </Dialog>

            {/* ******* 2ND API COLUMN HEADER FILTER MENU ******* */}
            <HeaderFilterMenu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor) && Boolean(activeHeader)}
                onClose={() => {
                    setFilterAnchor(null);
                    setActiveHeader(null);
                }}
                columnLabel={itemColumns.find((c) => c.key === activeHeader)?.label || activeHeader || undefined}
                options={filterOptions}
                selectedValues={activeHeader ? columnFilters[activeHeader] : undefined}
                onFilterChange={(selected: string[] | undefined) => {
                    if (activeHeader) {
                        setColumnFilters((prev) => {
                            const copy = { ...prev };
                            if (selected === undefined) {
                                delete copy[activeHeader];
                            } else {
                                copy[activeHeader] = selected;
                            }
                            return copy;
                        });
                    }
                }}
            />
        </Box>
    );
};

export default OverallItemwise;
