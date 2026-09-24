import React, { useEffect, useMemo, useState } from "react";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
import {
    Box,
    IconButton,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Paper,
    CircularProgress,
    Dialog,
    DialogTitle,
    DialogContent,
    Typography,
    Tooltip,
    Chip,
} from "@mui/material";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowRightIcon from "@mui/icons-material/KeyboardArrowRight";
import LayersIcon from "@mui/icons-material/Layers";
import CloseIcon from "@mui/icons-material/Close";
import dayjs from "dayjs";
import AppLayout, { useToggleMode } from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import { mapForExport } from "../../utils/exportMapper";
import {
    itemwisestockreportservice,
    godownwisestockreportservice,
    stockGroupingService,
    StockGroupConfig,
    stockWiseReport,
    godownStockBadgeService,
    GodownStockBadgeItem,
    itemwiseStockBadgeService,
    ItemwiseStockBadgeItem,
} from "../../services/stockWiseReport.service";


/* ================= UTIL ================= */

const sum = (rows: stockWiseReport[], key: keyof stockWiseReport) =>
    rows.reduce((s, r) => s + Number(r[key] || 0), 0);

/* ================= COMPONENT ================= */

type Level2Meta = {
    columnName: string;
    type: number;
};

const StockInHandReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const { toggleMode, setToggleMode } = useToggleMode();
    const isExpanded = toggleMode === "Expanded";
    const [rawData, setRawData] = useState<stockWiseReport[]>([]);
    const [groupConfig, setGroupConfig] = useState<StockGroupConfig[]>([]);
    const [expanded, setExpanded] = useState<Record<string, boolean>>({});
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    /* ===== BATCH MODAL STATES ===== */
    const [batchModalOpen, setBatchModalOpen] = useState(false);
    const [batchLoading, setBatchLoading] = useState(false);
    const [batchData, setBatchData] = useState<(GodownStockBadgeItem | ItemwiseStockBadgeItem)[]>([]);
    const [selectedBatchItem, setSelectedBatchItem] = useState<stockWiseReport | null>(null);

    /* ===== FILTER STATES ===== */

    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [filterLevels, setFilterLevels] = useState<Record<number, any[]>>({});
    const [selectedFilters, setSelectedFilters] = useState<Record<string, any>>({});
    const [loading, setLoading] = useState(false);

    const [qtyMode, setQtyMode] = useState<"qty" | "actQty">("qty");

    const qtyKeys = useMemo(() => {
        if (isExpanded) {
            if (qtyMode === "actQty") {
                return {
                    opening: "OB_Act_Qty" as keyof stockWiseReport,
                    in: "IN_Act_Qty" as keyof stockWiseReport,
                    out: "OUT_Act_Qty" as keyof stockWiseReport,
                    closing: "Act_Bal_Qty" as keyof stockWiseReport,
                };
            }
            return {
                opening: "OB_Bal_Qty" as keyof stockWiseReport,
                in: "IN_Qty" as keyof stockWiseReport,
                out: "OUT_Qty" as keyof stockWiseReport,
                closing: "Bal_Qty" as keyof stockWiseReport,
            };
        }
        if (qtyMode === "actQty") {
            return {
                opening: "OB_Act_Qty" as keyof stockWiseReport,
                in: "Pur_Act_Qty" as keyof stockWiseReport,
                out: "Sal_Act_Qty" as keyof stockWiseReport,
                closing: "Bal_Act_Qty" as keyof stockWiseReport,
            };
        }
        return {
            opening: "OB_Bal_Qty" as keyof stockWiseReport,
            in: "Pur_Qty" as keyof stockWiseReport,
            out: "Sal_Qty" as keyof stockWiseReport,
            closing: "Bal_Qty" as keyof stockWiseReport,
        };
    }, [qtyMode, isExpanded]);


    /* ===== LEVEL 2 META (FROM CONFIG) ===== */

    const [level2Meta, setLevel2Meta] = useState<Level2Meta[]>([]);
    const [level2TypeOrder, setLevel2TypeOrder] = useState<number[]>([]);
    const [selectedLevel2ByType, setSelectedLevel2ByType] =
        useState<Record<number, string>>({});
    const [stockFilter, setStockFilter] = useState<"hasValues" | "zero" | "all">("hasValues");

    const [modeState, setModeState] = useState<{
        Abstract: any;
        Expanded: any;
    }>({
        Abstract: {},
        Expanded: {},
    });
    /* ================= GROUP CONFIG ================= */

    useEffect(() => {
        const originalOverflow = document.body.style.overflow;
        const originalHeight = document.body.style.height;

        // ✅ Enable page scroll for this screen
        document.body.style.overflow = "auto";
        document.body.style.height = "auto";

        return () => {
            // 🔁 Restore when leaving the page
            document.body.style.overflow = originalOverflow;
            document.body.style.height = originalHeight;
        };
    }, []);


    useEffect(() => {
        const reportName =
            isExpanded ? "StockInhand-Godown" : "StockInhand";

        const loadConfig = async () => {
            try {
                setLoading(true);

                const res =
                    await stockGroupingService.getGroupingConfig(reportName);

                const cfg = res.data.data || [];

                const groups = cfg
                    .filter((g) => g.isGroupFilter)
                    .sort(
                        (a, b) =>
                            (a.Level_Id ?? 0) - (b.Level_Id ?? 0)
                    );

                setGroupConfig(groups);

                const filtersOnly = cfg.filter(
                    (g) => g.isGroupFilter === false
                );

                const grouped: Record<number, any[]> = {};

                filtersOnly.forEach((f) => {
                    const lvl = f.FilterLevel || 1;

                    if (!grouped[lvl]) grouped[lvl] = [];

                    grouped[lvl].push(f);
                });

                setFilterLevels(grouped);

                const lvl2 = cfg.filter(
                    (g) =>
                        g.FilterLevel === 2 &&
                        g.isGroupFilter === false
                );

                const meta = lvl2
                    .filter((l) => l.filterType)
                    .map((l) => ({
                        columnName: l.columnName,
                        type: Number(l.filterType),
                    }));

                setLevel2Meta(meta);

                setLevel2TypeOrder(
                    Array.from(
                        new Set(meta.map((m) => m.type))
                    ).sort((a, b) => a - b)
                );

                setSelectedLevel2ByType({});

            } catch (err) {
                console.error(err);
            } finally {
                setLoading(false);
            }
        };

        loadConfig();
    }, [toggleMode]);

    useEffect(() => {
        if (Object.keys(filterLevels).length > 0) {
            loadData();
        }
    }, [filterLevels]);

    useEffect(() => {
        setModeState(prev => ({
            ...prev,
            [toggleMode]: {
                rawData,
                expanded,
                page,
                selectedFilters,
                selectedLevel2ByType
            }
        }));
    }, [rawData, expanded, page, selectedFilters, selectedLevel2ByType, toggleMode]);

    useEffect(() => {
        const state = modeState[toggleMode];

        if (state) {
            setExpanded(state.expanded || {});
            setPage(state.page || 1);
            setSelectedFilters(state.selectedFilters || {});
            setSelectedLevel2ByType(state.selectedLevel2ByType || {});
        }
    }, [toggleMode]);

    const loadData = React.useCallback(async () => {
        try {
            setLoading(true);

            const payload: any = {
                Fromdate: fromDate,
                Todate: toDate,
            };

            /* =========================
               ABSTRACT MODE
            ========================= */
            if (!isExpanded) {
                Object.keys(selectedFilters).forEach((col) => {
                    const rawValue = selectedFilters[col];
                    const value = Array.isArray(rawValue) ? rawValue[0] : rawValue;
                    if (value === undefined || value === null || value === "") return;

                    const filter = Object.values(filterLevels)
                        .flat()
                        .find((f: any) => f.columnName === col);

                    const label = filter?.options?.find(
                        (o: any) => String(o.value) === String(value)
                    )?.label;

                    payload[col] = label || value;
                });

                const res =
                    await itemwisestockreportservice.getItemwiseReports(payload);

                setRawData(res.data?.data || []);
            }

            else {
                const godownFilter = (filterLevels[1] || []).find(
                    (f: any) =>
                        f.columnName === "Godown_Id" ||
                        f.columnName === "Godown_Name"
                );

                const rawGodownId =
                    selectedFilters["Godown_Id"] ||
                    selectedFilters["Godown_Name"] ||
                    (godownFilter ? selectedFilters[godownFilter.columnName] : undefined);

                if (!rawGodownId) {
                    setRawData([]);
                    setExpanded({});
                    setPage(1);
                    return;
                }

                const godownId = Array.isArray(rawGodownId) ? rawGodownId[0] : rawGodownId;
                if (!godownId) {
                    setRawData([]);
                    setExpanded({});
                    setPage(1);
                    return;
                }

                payload.Godown_Id = godownId;

                const sortedL1Filters = [...(filterLevels[1] || [])]
                    .filter((f: any) => f.columnName !== "Godown_Id" && f.columnName !== "Godown_Name")
                    .sort((a, b) => Number(a.filterType || 0) - Number(b.filterType || 0));

                const getFilterValuesString = (filter: any) => {
                    if (!filter) return undefined;
                    const values = selectedFilters[filter.columnName];
                    if (!values) return undefined;

                    const valArray = Array.isArray(values) ? values : [values];
                    if (valArray.length === 0) return undefined;

                    const labels = valArray.map(val => {
                        const opt = filter.options?.find((o: any) => String(o.value) === String(val));
                        return opt ? opt.label : String(val);
                    });
                    return labels.filter(Boolean).join(",");
                };

                payload.filter1 = getFilterValuesString(sortedL1Filters[0]);
                payload.filter2 = getFilterValuesString(sortedL1Filters[1]);
                payload.filter3 = getFilterValuesString(sortedL1Filters[2]);

                const res =
                    await godownwisestockreportservice.getGodownwiseReports(
                        payload
                    );

                setRawData(res.data?.data || []);
            }

            setExpanded({});
            setPage(1);

        } catch (err) {
            console.error("Stock report load error:", err);
            setRawData([]);
        } finally {
            setLoading(false);
        }
    }, [
        isExpanded,
        selectedFilters,
        fromDate,
        toDate,
        filterLevels
    ]);

    useEffect(() => {
        setExpanded({});
        setPage(1);
    }, [toggleMode]);

    useEffect(() => {
        const saved = sessionStorage.getItem("stockInHandState");

        if (!saved) return;
        const state = JSON.parse(saved);

        setRawData(state.rawData || []);
        setExpanded(state.expanded || {});
        setPage(state.page || 1);
        setSelectedFilters(state.selectedFilters || {});
        setSelectedLevel2ByType(state.selectedLevel2ByType || {});
        setToggleMode(state.toggleMode || "Abstract");
        setFromDate(state.fromDate || today);
        setToDate(state.toDate || today);
        setQtyMode(state.qtyMode || "qty");

        sessionStorage.removeItem("stockInHandState");

    }, []);

    /* ================= LEVEL 2 FILTER ================= */

    const data = useMemo(() => {
        let filtered = rawData;

        // LEVEL 1 FILTER
        // ✅ DYNAMIC FILTERS
        Object.keys(selectedFilters).forEach((col) => {
            const rawValue = selectedFilters[col];
            if (col === "Godown_Id" || col === "Godown_Name") return;

            const filter = Object.values(filterLevels)
                .flat()
                .find((f: any) => f.columnName === col);

            if (Array.isArray(rawValue)) {
                if (rawValue.length > 0) {
                    const labels = rawValue.map(val => {
                        const opt = filter?.options?.find((o: any) => String(o.value) === String(val));
                        return opt ? opt.label : String(val);
                    });
                    filtered = filtered.filter(
                        (r) => labels.some(lbl => String(r[col] || "").trim().toLowerCase() === String(lbl).trim().toLowerCase())
                    );
                }
            } else if (rawValue !== undefined && rawValue !== null && rawValue !== "") {
                const label = filter?.options?.find(
                    (opt: any) => String(opt.value) === String(rawValue)
                )?.label || (typeof rawValue === "string" ? rawValue : undefined);

                if (label) {
                    filtered = filtered.filter(
                        (r) =>
                            String(r[col] || "").trim().toLowerCase() === String(label).trim().toLowerCase() ||
                            String(r[col] || "").trim().toLowerCase() === String(rawValue).trim().toLowerCase()
                    );
                }
            }
        });

        // LEVEL 2 FILTER
        level2TypeOrder.forEach(type => {
            const selected = selectedLevel2ByType[type];
            if (!selected) return;

            const meta = level2Meta.find(m => m.type === type);
            if (!meta) return;

            filtered = filtered.filter(
                r => String(r[meta.columnName]) === String(selected)
            );
        });

        // ✅ ✅ ADD THIS BLOCK (IMPORTANT)
        filtered = filtered.filter((r) => {
            const ob = Number(r[qtyKeys.opening]) || 0;
            const input = Number(r[qtyKeys.in]) || 0;
            const out = Number(r[qtyKeys.out]) || 0;
            const cls = Number(r[qtyKeys.closing]) || 0;

            const hasValue = ob !== 0 || input !== 0 || out !== 0 || cls !== 0;
            const isZero = ob === 0 && input === 0 && out === 0 && cls === 0;

            if (stockFilter === "hasValues") return hasValue;
            if (stockFilter === "zero") return isZero;

            return true; // "all"
        });

        return filtered;
    }, [
        rawData,
        selectedFilters,
        level2TypeOrder,
        level2Meta,
        selectedLevel2ByType,
        stockFilter,
        qtyKeys
    ]);

    const {
        sortConfig: numSortConfig,
        rangeFilter: numRangeFilter,
        setRangeFilter: setNumRangeFilter,
        filterAnchor: numFilterAnchor,
        setFilterAnchor: setNumFilterAnchor,
        activeHeader: numActiveHeader,
        handleSort: handleNumSort,
        openFilter: openNumFilter,
        filteredAndSortedData: numFilteredAndSortedRows,
        getMinMax,
        clearRangeFilter,
    } = useNumericalFilter(data, ["OB_Bal_Qty", "Pur_Qty", "Sal_Qty", "Bal_Qty", "OB_Act_Qty", "Pur_Act_Qty", "Sal_Act_Qty", "Bal_Act_Qty", "IN_Qty", "OUT_Qty", "IN_Act_Qty", "OUT_Act_Qty", "Act_Bal_Qty"]);

    const handleTransactionClick = (
        row: stockWiseReport,
        mode: "ABSTRACT" | "EXPANDED"
    ) => {

        const path =
            mode === "EXPANDED"
                ? "/stockinhand/godown-item-transaction"
                : "/stockinhand/item-transaction";

        // ✅ Save report state
        sessionStorage.setItem(
            "stockInHandState",
            JSON.stringify({
                expanded,
                page,
                selectedFilters,
                selectedLevel2ByType,
                toggleMode,
                fromDate,
                toDate,
                qtyMode
            })
        );

        const params = new URLSearchParams({
            ProductId: String(row.Product_Id ?? ""),
            productName: row.stock_item_name ?? "",
            fromDate,
            toDate,
            Godown_Id: mode === "EXPANDED" ? String(row.Godown_Id ?? "") : "",
            godownName: mode === "EXPANDED" ? String(row.Godown_Name ?? "") : ""
        });

        const url = `${window.location.origin}${path}?${params.toString()}`;

        // ✅ open new tab
        window.open(url, "_blank", "noopener,noreferrer");
    };

    useEffect(() => {
        setExpanded({});
        setPage(1);
    }, [selectedLevel2ByType]);

    /* ================= GROUPING ================= */

    const buildGroups = (rows: stockWiseReport[], level: number): any[] => {
        const cfg = groupConfig[level];
        if (!cfg) return [];

        const map: Record<string, stockWiseReport[]> = {};
        rows.forEach((r: any) => {
            const key = r[cfg.columnName] || "Others";
            map[key] ||= [];
            map[key].push(r);
        });

        return Object.entries(map).map(([key, children]) => ({
            key,
            rows: children,
            level,
            children:
                level + 1 < groupConfig.length
                    ? buildGroups(children, level + 1)
                    : [],
        }));
    };

    // ✅ BASE DATA FOR LEVEL-2 CHIPS (MOBILE PARITY)
    const level1FilteredData = useMemo(() => {
        let filtered = rawData;

        Object.keys(selectedFilters).forEach((col) => {
            const rawValue = selectedFilters[col];
            if (col === "Godown_Id" || col === "Godown_Name") return;

            const filter = Object.values(filterLevels)
                .flat()
                .find((f: any) => f.columnName === col);

            if (Array.isArray(rawValue)) {
                if (rawValue.length > 0) {
                    const labels = rawValue.map(val => {
                        const opt = filter?.options?.find((o: any) => String(o.value) === String(val));
                        return opt ? opt.label : String(val);
                    });
                    filtered = filtered.filter(
                        (r) => labels.some(lbl => String(r[col] || "").trim().toLowerCase() === String(lbl).trim().toLowerCase())
                    );
                }
            } else if (rawValue !== undefined && rawValue !== null && rawValue !== "") {
                const label = filter?.options?.find(
                    (opt: any) => String(opt.value) === String(rawValue)
                )?.label || (typeof rawValue === "string" ? rawValue : undefined);

                if (label) {
                    filtered = filtered.filter(
                        (r) =>
                            String(r[col] || "").trim().toLowerCase() === String(label).trim().toLowerCase() ||
                            String(r[col] || "").trim().toLowerCase() === String(rawValue).trim().toLowerCase()
                    );
                }
            }
        });

        return filtered;
    }, [rawData, selectedFilters, filterLevels]);


    /* ===== GODOWN FIRST (EXPANDED MODE) ===== */
    const computeLevel2Values = (
        columnName: string,
        parent?: { column: string; value: string }
    ) => {
        const map = new Map<string, number>();

        level1FilteredData.forEach((r: any) => {
            if (parent) {
                if (String(r[parent.column]) !== parent.value) return;
            }

            const v = r[columnName];
            if (!v) return;

            const qty = Number(r[qtyKeys.closing] || 0);
            map.set(String(v), (map.get(String(v)) || 0) + qty);
        });

        return Array.from(map.entries())
            .map(([value, total]) => ({ value, total }))
            .sort((a, b) => b.total - a.total);
    };


    const finalGroups = useMemo(() => {
        if (!isExpanded) return buildGroups(numFilteredAndSortedRows, 0);

        const map: Record<string, stockWiseReport[]> = {};
        numFilteredAndSortedRows.forEach((r: any) => {
            const g = r.Godown_Name || "Unknown";
            map[g] ||= [];
            map[g].push(r);
        });

        return Object.entries(map).map(([key, rows]) => ({
            key,
            rows,
            level: -1,
            children: buildGroups(rows, 0),
        }));
    }, [numFilteredAndSortedRows, isExpanded, groupConfig]);

    const hasGrouping = groupConfig.length > 0;

    const formatQty = (value: any) =>
        Number(value || 0).toFixed(2);



    const extractBagKg = (row: any): number | null => {
        if (!row?.Bag) return null;

        const kg = parseFloat(String(row.Bag).toLowerCase().replace("kg", "").trim());
        return isNaN(kg) || kg <= 0 ? null : kg;
    };

    const formatBagCount = (value: number, decimals = 2) =>
        Number(value).toFixed(decimals);

    const formatQtyWithBag = (qty: string | number, row: any) => {
        const q = Number(qty || 0);
        const bagKg = extractBagKg(row);

        if (!bagKg) {
            return q.toFixed(2);
        }

        const bags = q / bagKg;

        return `${q.toFixed(2)} (${formatBagCount(bags)})`;
    };

    const getTotalBagCount = (rows: any[], qtyKey: keyof stockWiseReport) => {
        let totalBags = 0;
        let hasBag = false;

        rows.forEach(r => {
            const qty = Number(r[qtyKey] || 0);
            const bagKg = extractBagKg(r);

            if (bagKg && bagKg > 0) {
                totalBags += qty / bagKg;
                hasBag = true;
            }
        });

        if (!hasBag) return null;

        return Number(totalBags.toFixed(2));
    };


    const formatTotalQtyWithBag = (
        qty: number,
        rows: stockWiseReport[],
        qtyKey: keyof stockWiseReport
    ) => {
        const q = Number(qty || 0);
        const bags = getTotalBagCount(rows, qtyKey);

        if (!bags) {
            return q.toFixed(2);
        }

        return `${q.toFixed(2)} (${bags})`;
    };

    /* ================= BATCH (GODOWN STOCK BADGE) HELPERS ================= */

    const handleOpenBatchModal = async (item: stockWiseReport) => {
        setSelectedBatchItem(item);
        setBatchModalOpen(true);
        setBatchLoading(true);
        setBatchData([]);

        try {
            if (isExpanded) {
                const godownFilter = (filterLevels[1] || []).find(
                    (f: any) =>
                        f.columnName === "Godown_Id" ||
                        f.columnName === "Godown_Name"
                );
                const rawGodownId =
                    item.Godown_Id ||
                    selectedFilters["Godown_Id"] ||
                    selectedFilters["Godown_Name"] ||
                    (godownFilter ? selectedFilters[godownFilter.columnName] : undefined) ||
                    1;

                const godownId = Array.isArray(rawGodownId) ? rawGodownId[0] : rawGodownId;
                const itemId = item.Product_Id || item.Item_Id || "";

                const res = await godownStockBadgeService.getGodownStockBadge({
                    Fromdate: fromDate,
                    Todate: toDate,
                    Godown_Id: godownId,
                    Item_Id: itemId,
                });

                setBatchData(res.data?.data || []);
            } else {
                const res = await itemwiseStockBadgeService.getItemwiseStockBadge({
                    Fromdate: fromDate,
                    Todate: toDate,
                    Item_Name: item.stock_item_name,
                    stock_item_name: item.stock_item_name,
                    Product_Id: item.Product_Id,
                    Item_Id: item.Product_Id || item.Item_Id,
                });

                const rawBatches = res.data?.data || [];
                const targetName = String(item.stock_item_name || "").trim().toLowerCase();
                const targetId = String(item.Product_Id || item.Item_Id || "").trim();

                const filtered = rawBatches.filter((b: any) => {
                    if (targetName && String(b.stock_item_name || "").trim().toLowerCase() === targetName) {
                        return true;
                    }
                    if (targetId && String(b.Product_Id || "").trim() === targetId) {
                        return true;
                    }
                    return false;
                });

                setBatchData(filtered);
            }
        } catch (err) {
            console.error("Failed to load batch data:", err);
            setBatchData([]);
        } finally {
            setBatchLoading(false);
        }
    };

    const getBatchQty = (item: GodownStockBadgeItem | ItemwiseStockBadgeItem, type: "OB" | "IN" | "OUT" | "CLS") => {
        if (qtyMode === "actQty") {
            switch (type) {
                case "OB": return Number(item.OB_Act_Qty || 0);
                case "IN": return Number(item.IN_Act_Qty || 0);
                case "OUT": return Number(item.OUT_Act_Qty || 0);
                case "CLS": return Number(item.Act_Bal_Qty || 0);
            }
        }
        switch (type) {
            case "OB": return Number(item.OB_Bal_Qty || 0);
            case "IN": return Number(item.IN_Qty || 0);
            case "OUT": return Number(item.OUT_Qty || 0);
            case "CLS": return Number(item.Bal_Qty || 0);
        }
    };

    const formatBatchQtyWithBag = (qty: number, itemRow: stockWiseReport | null) => {
        const q = Number(qty || 0);
        const bagKg = extractBagKg(itemRow);
        if (!bagKg) {
            return q.toFixed(2);
        }
        const bags = q / bagKg;
        return `${q.toFixed(2)} (${formatBagCount(bags)})`;
    };


    const flattenGroupsForExport = (groups: any[], parentKeys: Record<string, string> = {}, isExpandedMode = isExpanded): any[] => {
        const result: any[] = [];

        groups.forEach(g => {
            const keys = { ...parentKeys };

            if (isExpandedMode && g.level === -1) {
                keys["Godown"] = g.key;
            } else if (g.level >= 0) {
                keys[`Group ${g.level + 1}`] = g.key;
            }

            if (g.children?.length) {
                result.push(...flattenGroupsForExport(g.children, keys, isExpandedMode));
            } else {
                g.rows.forEach((r: stockWiseReport) => {
                    result.push({
                        ...keys,
                        Item: r.stock_item_name,
                        Opening: r[qtyKeys.opening],
                        In: r[qtyKeys.in],
                        Out: r[qtyKeys.out],
                        Closing: r[qtyKeys.closing],
                    });
                });
            }
        });

        return result;
    };

    const getExportColumns = () => {
        const groupCols = groupConfig.map((_, i) => ({
            header: `Group ${i + 1}`,
            key: `Group ${i + 1}`,
        }));

        const baseCols = [
            { header: "Item", key: "Item" },
            { header: "Opening", key: "Opening" },
            { header: "In", key: "In" },
            { header: "Out", key: "Out" },
            { header: "Closing", key: "Closing" },
        ];

        if (isExpanded) {
            return [
                { header: "Godown", key: "Godown" },
                ...groupCols,
                ...baseCols,
            ];
        }

        return [...groupCols, ...baseCols];
    };

    const handleExportPDF = () => {
        const rows = flattenGroupsForExport(finalGroups);
        const columns = getExportColumns();

        const { headers, data } = mapForExport(columns, rows);

        exportToPDF(
            `Stock in Hand (${toggleMode})`,
            headers,
            data
        );
    };

    const handleExportExcel = () => {
        const rows = flattenGroupsForExport(finalGroups);
        const columns = getExportColumns();

        const { headers, data } = mapForExport(columns, rows);

        exportToExcel(
            `Stock in Hand (${toggleMode})`,
            headers,
            data
        );
    };


    /* ================= ITEM TABLE ================= */

    const paginated = (rows: stockWiseReport[]) =>
        rows.slice((page - 1) * rowsPerPage, page * rowsPerPage);

    /* ================= TOTALS ================= */

    const totalOpening = useMemo(() => sum(numFilteredAndSortedRows, qtyKeys.opening), [numFilteredAndSortedRows, qtyKeys.opening]);
    const totalIn = useMemo(() => sum(numFilteredAndSortedRows, qtyKeys.in), [numFilteredAndSortedRows, qtyKeys.in]);
    const totalOut = useMemo(() => sum(numFilteredAndSortedRows, qtyKeys.out), [numFilteredAndSortedRows, qtyKeys.out]);
    const totalClosing = useMemo(() => sum(numFilteredAndSortedRows, qtyKeys.closing), [numFilteredAndSortedRows, qtyKeys.closing]);

    /* ================= ITEM ROWS ================= */

    const renderItemRows = (items: stockWiseReport[], depth = 0) =>
        items.map((r, i) => (
            <TableRow
                key={`${r.Product_Id || r.Item_Id || i}-${i}`}
                hover
                sx={{
                    background: "#FFFFFF",
                    "&:hover": { bgcolor: "#EFF6FF" },
                    borderBottom: "1px solid #F1F5F9",
                }}
            >
                <TableCell width={44} sx={{ py: 0.6, textAlign: "center", color: "#94A3B8", fontSize: "0.75rem" }}>
                    {i + 1}
                </TableCell>

                <TableCell sx={{ py: 0.6 }}>
                    <Box sx={{ display: "flex", alignItems: "center", pl: depth * 2.5 + (depth > 0 ? 1 : 0), gap: 0.8, flexWrap: "wrap" }}>
                        <span
                            style={{ cursor: "pointer", color: "#1D4ED8", fontWeight: 600 }}
                            onClick={(e) => {
                                e.stopPropagation();
                                handleTransactionClick(
                                    r,
                                    isExpanded ? "EXPANDED" : "ABSTRACT"
                                );
                            }}
                            onMouseEnter={(e) => (e.currentTarget.style.textDecoration = "underline")}
                            onMouseLeave={(e) => (e.currentTarget.style.textDecoration = "none")}
                        >
                            {r.stock_item_name}
                        </span>

                        <Tooltip title="View Batch Stock Details" arrow placement="top">
                            <IconButton
                                size="small"
                                sx={{
                                    p: "3px",
                                    color: "#1E3A8A",
                                    backgroundColor: "#EFF6FF",
                                    border: "1px solid #BFDBFE",
                                    borderRadius: "6px",
                                    transition: "all 0.15s ease-in-out",
                                    "&:hover": {
                                        backgroundColor: "#DBEAFE",
                                        borderColor: "#93C5FD",
                                        transform: "scale(1.08)",
                                    },
                                }}
                                onClick={(e) => {
                                    e.stopPropagation();
                                    handleOpenBatchModal(r);
                                }}
                            >
                                <LayersIcon sx={{ fontSize: 16 }} />
                            </IconButton>
                        </Tooltip>
                    </Box>
                </TableCell>

                <TableCell align="right" sx={{ py: 0.6 }}>
                    {formatQtyWithBag(r[qtyKeys.opening] ?? 0, r)}
                </TableCell>

                <TableCell align="right" sx={{ py: 0.6 }}>
                    {formatQtyWithBag(r[qtyKeys.in] ?? 0, r)}
                </TableCell>

                <TableCell align="right" sx={{ py: 0.6 }}>
                    {formatQtyWithBag(r[qtyKeys.out] ?? 0, r)}
                </TableCell>

                <TableCell align="right" sx={{ py: 0.6 }}>
                    {formatQtyWithBag(r[qtyKeys.closing] ?? 0, r)}
                </TableCell>
            </TableRow>
        ));


    /* ================= GROUP ROWS ================= */

    const renderGroups = (groups: any[], depth = 0): React.ReactNode =>
        groups.map((g) => {
            const id = `${g.level}-${g.key}-${g.rows.length}`;
            const open = expanded[id];

            const bgColors = ["#F8FAFC", "#FFFFFF", "#F8FAFC"];
            const rowBg = depth === 0 ? "#F1F5F9" : bgColors[depth % bgColors.length];

            const grpOpening = sum(g.rows, qtyKeys.opening);
            const grpIn = sum(g.rows, qtyKeys.in);
            const grpOut = sum(g.rows, qtyKeys.out);
            const grpClosing = sum(g.rows, qtyKeys.closing);

            return (
                <React.Fragment key={id}>
                    <TableRow
                        sx={{
                            background: rowBg,
                            cursor: "pointer",
                            "&:hover": { bgcolor: "#E2E8F0" },
                            borderBottom: "1px solid #E2E8F0",
                        }}
                        onClick={() => setExpanded((p) => ({ ...p, [id]: !p[id] }))}
                    >
                        <TableCell width={44} sx={{ py: 0.8, textAlign: "center" }}>
                            {depth === 0 ? (
                                <IconButton
                                    size="small"
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        setExpanded((p) => ({ ...p, [id]: !p[id] }));
                                    }}
                                >
                                    {open ? <KeyboardArrowDownIcon fontSize="small" /> : <KeyboardArrowRightIcon fontSize="small" />}
                                </IconButton>
                            ) : null}
                        </TableCell>

                        <TableCell sx={{ py: 0.8 }}>
                            <Box sx={{ display: "flex", alignItems: "center", pl: depth * 2.5, gap: 0.6 }}>
                                {depth > 0 && (
                                    <IconButton
                                        size="small"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            setExpanded((p) => ({ ...p, [id]: !p[id] }));
                                        }}
                                        sx={{ p: 0.5 }}
                                    >
                                        {open ? <KeyboardArrowDownIcon fontSize="small" /> : <KeyboardArrowRightIcon fontSize="small" />}
                                    </IconButton>
                                )}
                                <span style={{ fontWeight: depth === 0 ? 700 : 600, color: "#1E293B" }}>
                                    {g.key}
                                </span>
                                <span style={{ color: "#64748B", fontWeight: 600, fontSize: "0.82rem" }}>
                                    ({g.rows.length})
                                </span>
                            </Box>
                        </TableCell>

                        <TableCell align="right" sx={{ fontWeight: 600, py: 0.8 }}>
                            {formatTotalQtyWithBag(grpOpening, g.rows, qtyKeys.opening)}
                        </TableCell>

                        <TableCell align="right" sx={{ fontWeight: 600, py: 0.8 }}>
                            {formatTotalQtyWithBag(grpIn, g.rows, qtyKeys.in)}
                        </TableCell>

                        <TableCell align="right" sx={{ fontWeight: 600, py: 0.8 }}>
                            {formatTotalQtyWithBag(grpOut, g.rows, qtyKeys.out)}
                        </TableCell>

                        <TableCell align="right" sx={{ fontWeight: 600, py: 0.8 }}>
                            {formatTotalQtyWithBag(grpClosing, g.rows, qtyKeys.closing)}
                        </TableCell>
                    </TableRow>

                    {open && (
                        g.children?.length
                            ? renderGroups(g.children, depth + 1)
                            : renderItemRows(g.rows, depth + 1)
                    )}
                </React.Fragment>
            );
        });

    /* ================= RENDER ================= */

    return (
        <>
            <PageHeader
                toggleMode={toggleMode}
                onToggleChange={setToggleMode}
                onExportPDF={handleExportPDF}
                onExportExcel={handleExportExcel}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen(prev => !prev)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                toDate={toDate}
                onFromDateChange={setFromDate}
                onToDateChange={setToDate}
                filterLevels={filterLevels}
                selectedFilters={selectedFilters}
                level1Multiple={false}
                onFilterChange={(col, val) => {
                    setSelectedFilters((prev) => {
                        const copy = { ...prev };
                        if (val === undefined || val === null || val === "") {
                            delete copy[col];
                        } else {
                            copy[col] = val;
                        }
                        return copy;
                    });
                }}
                stockFilter={stockFilter}
                onStockFilterChange={setStockFilter}
                showQtyModeFilter
                qtyModeValue={qtyMode}
                onQtyModeChange={setQtyMode}
                onApply={async () => {
                    await loadData();
                    setDrawerOpen(false);
                }}
            />

            <AppLayout fullWidth>
                {/* LEVEL 2 CHIPS */}
                <Box
                    sx={{
                        maxHeight: 180,
                        overflowY: "auto",
                        px: 1,
                        py: 1,
                        borderBottom: "1px solid #E5E7EB",
                        background: "#F8FAFC",
                    }}
                >
                    {level2TypeOrder.map((type, idx) => {
                        const meta = level2Meta.find(m => m.type === type);
                        if (!meta) return null;

                        let parent;
                        if (idx > 0) {
                            const parentType = level2TypeOrder[idx - 1];
                            const parentValue = selectedLevel2ByType[parentType];
                            if (!parentValue) return null;

                            const parentMeta = level2Meta.find(m => m.type === parentType);
                            if (parentMeta) {
                                parent = {
                                    column: parentMeta.columnName,
                                    value: parentValue,
                                };
                            }
                        }

                        const values = computeLevel2Values(meta.columnName, parent);
                        if (!values.length) return null;

                        const selected = selectedLevel2ByType[type];

                        return (
                            <Box
                                key={type}
                                sx={{
                                    display: "flex",
                                    flexWrap: "wrap",
                                    gap: 1,
                                    mb: 1,
                                }}
                            >
                                {/* ALL */}
                                <Box
                                    onClick={() => {
                                        setSelectedLevel2ByType(prev => {
                                            const copy = { ...prev };
                                            delete copy[type];
                                            level2TypeOrder.forEach(t => t > type && delete copy[t]);
                                            return copy;
                                        });
                                    }}
                                    sx={{
                                        px: 1.5,
                                        py: 0.5,
                                        borderRadius: "12px",
                                        cursor: "pointer",
                                        fontSize: "0.75rem",
                                        fontWeight: 600,
                                        background: !selected ? "#1E3A8A" : "#E5E7EB",
                                        color: !selected ? "#fff" : "#111",
                                    }}
                                >
                                    ALL
                                </Box>

                                {values.map(v => (
                                    <Box
                                        key={v.value}
                                        onClick={() => {
                                            setSelectedLevel2ByType(prev => {
                                                const copy = { ...prev, [type]: v.value };
                                                level2TypeOrder.forEach(t => t > type && delete copy[t]);
                                                return copy;
                                            });
                                        }}
                                        sx={{
                                            px: 1.5,
                                            py: 0.5,
                                            borderRadius: "16px",
                                            cursor: "pointer",
                                            fontSize: "0.75rem",
                                            fontWeight: 600,
                                            background:
                                                selected === v.value ? "#1E3A8A" : "#E5E7EB",
                                            color:
                                                selected === v.value ? "#fff" : "#111",
                                        }}
                                    >
                                        {v.value} ({formatQty(v.total)})
                                    </Box>
                                ))}
                            </Box>
                        );
                    })}
                </Box>

                <Paper
                    sx={{
                        mx: 1,
                        display: "flex",
                        flexDirection: "column",
                        maxHeight: "calc(100vh - 130px)",
                        overflow: "hidden",
                    }}
                >
                    <TableContainer sx={{ maxHeight: "calc(100vh - 130px)", overflow: "auto" }}>
                        {loading ? (
                            <Box
                                sx={{
                                    height: 300,
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    fontWeight: 600,
                                    color: "#64748B"
                                }}
                            >
                                <CircularProgress />
                            </Box>
                        ) : (
                            <Table size="small" stickyHeader>
                                <TableHead
                                    sx={{
                                        background: "#1E3A8A",
                                        position: "sticky",
                                        top: 0,
                                        zIndex: 3,
                                        "& .MuiTableCell-root": {
                                            bgcolor: "#1E3A8A !important",
                                            color: "#fff !important",
                                            fontWeight: 700,
                                            fontSize: "0.82rem",
                                            borderBottom: "1px solid #1E3A8A",
                                        },
                                        "& .MuiTableCell-stickyHeader": {
                                            bgcolor: "#1E3A8A !important",
                                            color: "#fff !important",
                                        },
                                    }}
                                >
                                    <TableRow>
                                        <TableCell sx={{ color: "#fff", width: 44, fontWeight: 700, p: 1, bgcolor: "#1E3A8A" }} />
                                        <TableCell sx={{ color: "#fff", fontWeight: 700, bgcolor: "#1E3A8A" }}>Name</TableCell>
                                        <TableCell sx={{ color: "#fff", fontWeight: 700, width: 140, bgcolor: "#1E3A8A" }} align="right">
                                            <SortableHeaderLabel
                                                label="OB"
                                                columnKey={qtyKeys.opening as string}
                                                sortConfig={numSortConfig}
                                                onSort={handleNumSort}
                                                onOpenFilter={(e) => openNumFilter(e, qtyKeys.opening as string)}
                                            />
                                        </TableCell>
                                        <TableCell sx={{ color: "#fff", fontWeight: 700, width: 140, bgcolor: "#1E3A8A" }} align="right">
                                            <SortableHeaderLabel
                                                label="IN"
                                                columnKey={qtyKeys.in as string}
                                                sortConfig={numSortConfig}
                                                onSort={handleNumSort}
                                                onOpenFilter={(e) => openNumFilter(e, qtyKeys.in as string)}
                                            />
                                        </TableCell>
                                        <TableCell sx={{ color: "#fff", fontWeight: 700, width: 140, bgcolor: "#1E3A8A" }} align="right">
                                            <SortableHeaderLabel
                                                label="OUT"
                                                columnKey={qtyKeys.out as string}
                                                sortConfig={numSortConfig}
                                                onSort={handleNumSort}
                                                onOpenFilter={(e) => openNumFilter(e, qtyKeys.out as string)}
                                            />
                                        </TableCell>
                                        <TableCell sx={{ color: "#fff", fontWeight: 700, width: 140, bgcolor: "#1E3A8A" }} align="right">
                                            <SortableHeaderLabel
                                                label="CLS"
                                                columnKey={qtyKeys.closing as string}
                                                sortConfig={numSortConfig}
                                                onSort={handleNumSort}
                                                onOpenFilter={(e) => openNumFilter(e, qtyKeys.closing as string)}
                                            />
                                        </TableCell>
                                    </TableRow>
                                </TableHead>

                                <TableBody>
                                    {/* ✅ TOTAL ROW — directly below header */}
                                    <TableRow sx={{ background: "#F1F5F9", fontWeight: 700, position: "sticky", top: 37, zIndex: 1 }}>
                                        <TableCell width={44} />
                                        <TableCell sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            TOTAL
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatTotalQtyWithBag(totalOpening, numFilteredAndSortedRows, qtyKeys.opening)}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatTotalQtyWithBag(totalIn, numFilteredAndSortedRows, qtyKeys.in)}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatTotalQtyWithBag(totalOut, numFilteredAndSortedRows, qtyKeys.out)}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatTotalQtyWithBag(totalClosing, numFilteredAndSortedRows, qtyKeys.closing)}
                                        </TableCell>
                                    </TableRow>

                                    {numFilteredAndSortedRows.length === 0 ? (
                                        <TableRow>
                                            <TableCell colSpan={6} align="center" sx={{ py: 4, color: "#64748B" }}>
                                                No records found
                                            </TableCell>
                                        </TableRow>
                                    ) : hasGrouping ? (
                                        renderGroups(finalGroups, 0)
                                    ) : (
                                        renderItemRows(paginated(numFilteredAndSortedRows), 0)
                                    )}
                                </TableBody>
                            </Table>
                        )}
                    </TableContainer>

                </Paper>

                <CommonPagination
                    totalRows={numFilteredAndSortedRows.length}
                    page={page}
                    rowsPerPage={rowsPerPage}
                    onPageChange={setPage}
                    onRowsPerPageChange={setRowsPerPage}
                />
            </AppLayout>

            <NumericalFilterMenu
                anchorEl={numFilterAnchor}
                open={Boolean(numFilterAnchor)}
                onClose={() => setNumFilterAnchor(null)}
                activeHeader={numActiveHeader}
                min={numActiveHeader ? getMinMax(numActiveHeader).min : 0}
                max={numActiveHeader ? getMinMax(numActiveHeader).max : 100}
                rangeFilter={numRangeFilter}
                onRangeChange={(key, range) => setNumRangeFilter(p => ({ ...p, [key]: range }))}
                onClear={clearRangeFilter}
            />

            {/* ================= BATCH DETAILS MODAL (EXPANDED SIDE) ================= */}
            <Dialog
                open={batchModalOpen}
                onClose={() => setBatchModalOpen(false)}
                maxWidth="lg"
                fullWidth
                PaperProps={{
                    sx: {
                        borderRadius: 2.5,
                        boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)",
                        overflow: "hidden",
                        maxWidth: "1100px",
                        width: "95%",
                    }
                }}
            >
                <DialogTitle
                    sx={{
                        m: 0,
                        px: 3,
                        py: 2,
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        bgcolor: "#1E3A8A",
                        color: "#fff",
                    }}
                >
                    <Box>
                        <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 0.5 }}>
                            <LayersIcon sx={{ fontSize: 22, color: "#93C5FD" }} />
                            <Typography variant="h6" component="div" fontWeight={700} sx={{ fontSize: "1.05rem" }}>
                                Batch Stock Details
                            </Typography>
                            <Chip
                                label={qtyMode === "actQty" ? "Actual Qty Mode" : "Normal Qty Mode"}
                                size="small"
                                sx={{
                                    bgcolor: qtyMode === "actQty" ? "#0284C7" : "#0D9488",
                                    color: "#fff",
                                    fontWeight: 600,
                                    fontSize: "0.7rem",
                                    height: 22,
                                }}
                            />
                        </Box>
                        <Typography variant="caption" sx={{ opacity: 0.9, fontWeight: 500, display: "block" }}>
                            <strong>Item:</strong> {selectedBatchItem?.stock_item_name || "—"} &nbsp;|&nbsp;{" "}
                            {isExpanded && (
                                <>
                                    <strong>Godown:</strong> {selectedBatchItem?.Godown_Name || (batchData[0] as GodownStockBadgeItem)?.Godown_Name || "—"} &nbsp;|&nbsp;{" "}
                                </>
                            )}
                            <strong>Period:</strong> {dayjs(fromDate).format("DD/MM/YYYY")} - {dayjs(toDate).format("DD/MM/YYYY")}
                        </Typography>
                    </Box>
                    <IconButton
                        aria-label="close"
                        onClick={() => setBatchModalOpen(false)}
                        sx={{
                            color: "#fff",
                            "&:hover": { bgcolor: "rgba(255,255,255,0.15)" }
                        }}
                    >
                        <CloseIcon />
                    </IconButton>
                </DialogTitle>

                <DialogContent dividers sx={{ p: 2.5, bgcolor: "#F8FAFC", overflowX: "hidden" }}>
                    {batchLoading ? (
                        <Box display="flex" flexDirection="column" justifyContent="center" alignItems="center" minHeight={200} gap={1.5}>
                            <CircularProgress size={36} sx={{ color: "#1E3A8A" }} />
                            <Typography variant="body2" sx={{ color: "#64748B", fontWeight: 500 }}>
                                Loading batch details...
                            </Typography>
                        </Box>
                    ) : batchData.length === 0 ? (
                        <Box display="flex" flexDirection="column" justifyContent="center" alignItems="center" minHeight={180} gap={1}>
                            <Typography variant="body1" sx={{ color: "#475569", fontWeight: 600 }}>
                                {isExpanded
                                    ? "No batch stock found for this item in this godown."
                                    : "No batch stock found for this item."}
                            </Typography>
                            <Typography variant="caption" sx={{ color: "#94A3B8" }}>
                                Try changing the date range in the Report Filter Drawer.
                            </Typography>
                        </Box>
                    ) : (
                        <TableContainer component={Paper} elevation={0} sx={{ border: "1px solid #CBD5E1", borderRadius: 2, maxHeight: 460, overflowX: "hidden" }}>
                            <Table size="small" stickyHeader>
                                <TableHead>
                                    <TableRow>
                                        <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, width: 60 }}>
                                            S.No
                                        </TableCell>
                                        <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600 }}>
                                            Batch No
                                        </TableCell>
                                        <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600 }}>
                                            OB
                                        </TableCell>
                                        <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600 }}>
                                            IN
                                        </TableCell>
                                        <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600 }}>
                                            OUT
                                        </TableCell>
                                        <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600 }}>
                                            CLS
                                        </TableCell>
                                    </TableRow>
                                </TableHead>
                                <TableBody>
                                    {/* TOTAL ROW */}
                                    <TableRow sx={{ background: "#F1F5F9", fontWeight: 700, position: "sticky", top: 37, zIndex: 1 }}>
                                        <TableCell colSpan={2} sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            TOTAL ({batchData.length} Batches)
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatBatchQtyWithBag(
                                                batchData.reduce((s, b) => s + getBatchQty(b, "OB"), 0),
                                                selectedBatchItem
                                            )}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatBatchQtyWithBag(
                                                batchData.reduce((s, b) => s + getBatchQty(b, "IN"), 0),
                                                selectedBatchItem
                                            )}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatBatchQtyWithBag(
                                                batchData.reduce((s, b) => s + getBatchQty(b, "OUT"), 0),
                                                selectedBatchItem
                                            )}
                                        </TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, color: "#0F172A" }}>
                                            {formatBatchQtyWithBag(
                                                batchData.reduce((s, b) => s + getBatchQty(b, "CLS"), 0),
                                                selectedBatchItem
                                            )}
                                        </TableCell>
                                    </TableRow>

                                    {/* BATCH ROWS */}
                                    {batchData.map((batch, idx) => (
                                        <TableRow
                                            key={idx}
                                            hover
                                            sx={{
                                                "&:nth-of-type(even)": { backgroundColor: "#FAFAFA" },
                                                "&:last-child td, &:last-child th": { border: 0 }
                                            }}
                                        >
                                            <TableCell sx={{ color: "#64748B", fontWeight: 500 }}>
                                                {idx + 1}
                                            </TableCell>
                                            <TableCell sx={{ fontWeight: 600, color: "#1E293B" }}>
                                                {batch.Batch_No || "—"}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 500 }}>
                                                {formatBatchQtyWithBag(getBatchQty(batch, "OB"), selectedBatchItem)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 500, color: getBatchQty(batch, "IN") > 0 ? "#16A34A" : "inherit" }}>
                                                {formatBatchQtyWithBag(getBatchQty(batch, "IN"), selectedBatchItem)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 500, color: getBatchQty(batch, "OUT") > 0 ? "#DC2626" : "inherit" }}>
                                                {formatBatchQtyWithBag(getBatchQty(batch, "OUT"), selectedBatchItem)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 600, color: "#1D4ED8" }}>
                                                {formatBatchQtyWithBag(getBatchQty(batch, "CLS"), selectedBatchItem)}
                                            </TableCell>
                                        </TableRow>
                                    ))}
                                </TableBody>
                            </Table>
                        </TableContainer>
                    )}
                </DialogContent>
            </Dialog>
        </>
    );
};

export default StockInHandReport;
