import React, { useEffect, useMemo, useState } from "react";
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
    CircularProgress
} from "@mui/material";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowRightIcon from "@mui/icons-material/KeyboardArrowRight";
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

    /* ===== FILTER STATES ===== */

    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [filterLevels, setFilterLevels] = useState<Record<number, any[]>>({});
    const [selectedFilters, setSelectedFilters] = useState<Record<string, any>>({});
    const [loading, setLoading] = useState(false);


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
                    const value = selectedFilters[col];

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

        setRawData(state.rawData || []); 41
        setExpanded(state.expanded || {});
        setPage(state.page || 1);
        setSelectedFilters(state.selectedFilters || {});
        setSelectedLevel2ByType(state.selectedLevel2ByType || {});
        setToggleMode(state.toggleMode || "Abstract");
        setFromDate(state.fromDate || today);
        setToDate(state.toDate || today);

        sessionStorage.removeItem("stockInHandState");

    }, []);

    /* ================= LEVEL 2 FILTER ================= */

    const data = useMemo(() => {
        let filtered = rawData;

        // LEVEL 1 FILTER
        // ✅ DYNAMIC FILTERS
        Object.keys(selectedFilters).forEach((col) => {
            const value = selectedFilters[col];
            if (col === "Godown_Id" || col === "Godown_Name") return;

            const filter = Object.values(filterLevels)
                .flat()
                .find((f: any) => f.columnName === col);

            if (Array.isArray(value)) {
                if (value.length > 0) {
                    const labels = value.map(val => {
                        const opt = filter?.options?.find((o: any) => String(o.value) === String(val));
                        return opt ? opt.label : String(val);
                    });
                    filtered = filtered.filter(
                        (r) => labels.some(lbl => String(r[col] || "").trim().toLowerCase() === String(lbl).trim().toLowerCase())
                    );
                }
            } else {
                const label = filter?.options?.find(
                    (opt: any) => String(opt.value) === String(value)
                )?.label;

                if (label) {
                    filtered = filtered.filter(
                        (r) => String(r[col] || "").trim().toLowerCase() === String(label).trim().toLowerCase()
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
            const ob = Number(r.OB_Bal_Qty) || 0;
            const input = Number(r.Pur_Qty) || 0;
            const out = Number(r.Sal_Qty) || 0;
            const cls = Number(r.Bal_Qty) || 0;

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
        stockFilter
    ]);

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
                toDate
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
            const value = selectedFilters[col];

            const filter = Object.values(filterLevels)
                .flat()
                .find((f: any) => f.columnName === col);

            const label = filter?.options?.find(
                (opt: any) => String(opt.value) === String(value)
            )?.label;

            if (label) {
                filtered = filtered.filter(
                    (r) => String(r[col]) === String(label)
                );
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

            const qty = Number(r.Bal_Qty || 0);
            map.set(String(v), (map.get(String(v)) || 0) + qty);
        });

        return Array.from(map.entries())
            .map(([value, total]) => ({ value, total }))
            .sort((a, b) => b.total - a.total);
    };


    const finalGroups = useMemo(() => {
        if (!isExpanded) return buildGroups(data, 0);

        const map: Record<string, stockWiseReport[]> = {};
        data.forEach((r: any) => {
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
    }, [data, isExpanded, groupConfig]);

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
                        Opening: r.OB_Bal_Qty,
                        In: r.Pur_Qty,
                        Out: r.Sal_Qty,
                        Closing: r.Bal_Qty,
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

    /* ================= ITEM TABLE ================= */

    const renderItemTable = (rows: stockWiseReport[]) => {
        const pageRows = paginated(rows);

        // ✅ TOTALS (full filtered rows, NOT paginated)
        const totalOpening = sum(rows, "OB_Bal_Qty");
        const totalIn = sum(rows, "Pur_Qty");
        const totalOut = sum(rows, "Sal_Qty");
        const totalClosing = sum(rows, "Bal_Qty");

        return (
            <Table size="small">
                <TableHead sx={{ background: "#1E3A8A" }}>
                    <TableRow>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }}>S.No</TableCell>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }}>Item</TableCell>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }} align="right">
                            Opening
                        </TableCell>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }} align="right">
                            In
                        </TableCell>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }} align="right">
                            Out
                        </TableCell>
                        <TableCell sx={{ color: "#fff", fontWeight: 600 }} align="right">
                            Closing
                        </TableCell>
                    </TableRow>
                </TableHead>

                <TableBody>
                    {/* ✅ TOTAL ROW — directly below header */}
                    <TableRow sx={{ background: "#F1F5F9", fontWeight: 700 }}>
                        <TableCell colSpan={2}>TOTAL</TableCell>

                        <TableCell align="right">
                            {formatTotalQtyWithBag(totalOpening, rows, "OB_Bal_Qty")}
                        </TableCell>

                        <TableCell align="right">
                            {formatTotalQtyWithBag(totalIn, rows, "Pur_Qty")}
                        </TableCell>

                        <TableCell align="right">
                            {formatTotalQtyWithBag(totalOut, rows, "Sal_Qty")}
                        </TableCell>

                        <TableCell align="right">
                            {formatTotalQtyWithBag(totalClosing, rows, "Bal_Qty")}
                        </TableCell>
                    </TableRow>

                    {/* ✅ ITEM ROWS */}
                    {pageRows.map((r, i) => (
                        <TableRow key={i}>
                            <TableCell>
                                {(page - 1) * rowsPerPage + i + 1}
                            </TableCell>
                            <TableCell
                                sx={{
                                    cursor: "pointer",
                                    color: "#1D4ED8",
                                    fontWeight: 600,
                                    "&:hover": { textDecoration: "none" }
                                }}
                                onClick={(e) => {
                                    e.stopPropagation();
                                    handleTransactionClick(
                                        r,
                                        isExpanded ? "EXPANDED" : "ABSTRACT"
                                    );
                                }}
                            >
                                {r.stock_item_name}
                            </TableCell>
                            <TableCell align="right">
                                {formatQtyWithBag(r.OB_Bal_Qty, r)}
                            </TableCell>
                            <TableCell align="right">
                                {formatQtyWithBag(r.Pur_Qty, r)}
                            </TableCell>
                            <TableCell align="right">
                                {formatQtyWithBag(r.Sal_Qty, r)}
                            </TableCell>
                            <TableCell align="right">
                                {formatQtyWithBag(r.Bal_Qty, r)}
                            </TableCell>
                        </TableRow>
                    ))}
                </TableBody>
            </Table>
        );
    };


    /* ================= GROUP ROWS ================= */

    const renderGroups = (groups: any[]) =>
        groups.map((g) => {
            const id = `${g.level}-${g.key}-${g.rows.length}`;
            const open = expanded[id];

            return (
                <React.Fragment key={id}>
                    <TableRow sx={{ background: "#F1F5F9" }}>
                        <TableCell width={40}>
                            <IconButton
                                size="small"
                                onClick={() =>
                                    setExpanded((p) => ({ ...p, [id]: !p[id] }))
                                }
                            >
                                {open ? <KeyboardArrowDownIcon /> : <KeyboardArrowRightIcon />}
                            </IconButton>
                        </TableCell>

                        <TableCell sx={{ fontWeight: 600 }}>{g.key}</TableCell>
                        <TableCell align="right">{g.rows.length}</TableCell>
                        <TableCell align="right">
                            {formatTotalQtyWithBag(
                                sum(g.rows, "Bal_Qty"),
                                g.rows,
                                "Bal_Qty"
                            )}
                        </TableCell>
                    </TableRow>

                    {open && (
                        <TableRow>
                            {/* colSpan should cover all columns including icon + totals */}
                            <TableCell colSpan={6} sx={{ pl: 4 }}>
                                {g.children.length
                                    ? (
                                        <Table size="small">
                                            <TableBody>{renderGroups(g.children)}</TableBody>
                                        </Table>
                                    )
                                    : renderItemTable(g.rows)}
                            </TableCell>
                        </TableRow>
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
                onFilterChange={(col, val) => {
                    setSelectedFilters((prev) => ({
                        ...prev,
                        [col]: val,
                    }));
                }}
                stockFilter={stockFilter}
                onStockFilterChange={setStockFilter}
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
                        overflow: "visible",
                    }}
                >
                    <TableContainer >
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
                        ) : hasGrouping ? (
                            <Table size="small">
                                <TableHead sx={{ background: "#1E3A8A" }}>
                                    <TableRow>
                                        <TableCell />
                                        <TableCell sx={{ color: "#fff" }}>Name</TableCell>
                                        <TableCell sx={{ color: "#fff" }} align="right">
                                            Items
                                        </TableCell>
                                        <TableCell sx={{ color: "#fff" }} align="right">
                                            Balance
                                        </TableCell>
                                    </TableRow>
                                </TableHead>

                                <TableBody>
                                    {renderGroups(finalGroups)}
                                </TableBody>
                            </Table>
                        ) : (
                            renderItemTable(data)
                        )}
                    </TableContainer>

                </Paper>

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

export default StockInHandReport;
