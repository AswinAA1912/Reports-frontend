import React, { useState, useMemo, useEffect } from "react";
import {
  AppBar,
  Toolbar,
  Box,
  Select,
  MenuItem,
  ToggleButton,
  ToggleButtonGroup,
  IconButton,
  Menu,
  Tooltip,
  Button,
  alpha,
  useTheme,
  useMediaQuery,
} from "@mui/material";
import MenuIcon from "@mui/icons-material/Menu";
import FileDownloadIcon from "@mui/icons-material/FileDownload";
import PowerSettingsNewIcon from "@mui/icons-material/PowerSettingsNew";
import ArrowDropDownIcon from "@mui/icons-material/ArrowDropDown";
import logo from "../assets/logo.png";
import { SettingsService } from "../services/reportSettings.services";
import { useNavigate, useLocation } from "react-router-dom";
import { useAuth, Company } from "../auth/authContext";
import { MenuService } from "../services/menus.service";

export type ToggleMode = "Abstract" | "Expanded";

export type Menus = {
  id: number;
  name: string;
  rUrl: string;
  is_active: number;
};

export interface Page {
  label: string;
  path: string;
}

interface PageHeaderProps {
  /** OPTIONAL TOGGLE */
  toggleMode?: ToggleMode;
  onToggleChange?: (mode: ToggleMode) => void;

  /** OPTIONAL EXPORT */
  onExportPDF?: () => void;
  onExportExcel?: () => void;

  /** ✅ OPTIONAL SETTINGS SLOT */
  settingsSlot?: React.ReactNode;
  infoSlot?: React.ReactNode;
  showPages?: boolean;
  onReportChange?: (report: any) => void;
  onQuickSave?: (parentName: string) => void;
}

export const PAGE_HEADER_HEIGHT = 40;
export const PAGE_HEADER_HEIGHT_MOBILE = 72;

const PageHeader: React.FC<PageHeaderProps> = ({
  toggleMode,
  onToggleChange,
  onExportPDF,
  onExportExcel,
  settingsSlot,
  infoSlot,
  showPages = true,
  onReportChange,
  onQuickSave
}) => {
  const navigate = useNavigate();
  const location = useLocation();

  const { user, companies, switchCompany, logout, isAutoLogin } = useAuth();
  const [anchorElCompany, setAnchorElCompany] = useState<null | HTMLElement>(null);
  const openCompanyMenu = Boolean(anchorElCompany);
  const [pages, setPages] = useState<Page[]>([]);
  const [routeMap, setRouteMap] = useState<Record<string, string>>({});
  const [anchorElExport, setAnchorElExport] = useState<null | HTMLElement>(null);
  const openExportMenu = Boolean(anchorElExport);
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down("sm"));
  const [templates, setTemplates] = useState<any[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState("");


  // Company name for display
  const companyName = useMemo(() => user?.Company_Name || "", [user]);

  useEffect(() => {
    const fetchMenus = async () => {
      try {
        const res = await MenuService.getMenus();
        const menus = res.data.data;

        const parentPages: Page[] = [];
        const subRouteMap: Record<string, string> = {};

        menus.forEach((menu: any) => {
          if (
            menu.menu_type === 1 &&
            menu.is_active === 3 &&
            menu.SubMenu?.length
          ) {
            menu.SubMenu.forEach((sub: any) => {
              if (sub.is_active === 3) {
                parentPages.push({
                  label: sub.name,
                  path: sub.rUrl,
                });

                if (sub.SubRoutes?.length) {
                  sub.SubRoutes.forEach((route: any) => {
                    if (route.is_active === 3) {
                      subRouteMap[route.rUrl] = sub.rUrl;
                    }
                  });
                }
              }
            });
          }
        });

        setPages(parentPages);
        setRouteMap(subRouteMap);
      } catch (err) {
        console.error("Menu fetch error", err);
      }
    };

    fetchMenus();
  }, []);

  // Handle company menu
  const handleCompanyClick = (event: React.MouseEvent<HTMLElement>) => {
    setAnchorElCompany(event.currentTarget);
  };

  const handleCompanyClose = async (company?: Company) => {
    setAnchorElCompany(null);
    if (company) {
      try {
        await switchCompany(company);
        window.location.reload();
      } catch (e) {
        // Switch failed, do not reload
      }
    }
  };

  const selectedPath = routeMap[location.pathname] || location.pathname;

  const currentPageLabel =
    pages.find((p) => p.path === selectedPath)?.label ||
    location.pathname.replace("/", "").replace(/-/g, " ").toUpperCase();

  useEffect(() => {
    if (!currentPageLabel) return;

    const loadTemplates = async () => {
      try {
        const res = await SettingsService.getReportsByParent(currentPageLabel);

        if (res.data.success) {
          setTemplates(res.data.data || []);
        }
      } catch (err) {
        console.error("Template load error", err);
      }
    };

    loadTemplates();
  }, [currentPageLabel]);

  return (
    <>
      <AppBar
        position="fixed"
        elevation={1}
        sx={{
          background: "#1E3A8A",
          height: isMobile ? PAGE_HEADER_HEIGHT_MOBILE : PAGE_HEADER_HEIGHT,
          top: 0,
          zIndex: (theme) => theme.zIndex.appBar,
        }}
      >
        <Toolbar
          disableGutters
          sx={{
            minHeight: `${PAGE_HEADER_HEIGHT}px !important`,
            height: PAGE_HEADER_HEIGHT,
            px: 1,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 1,
            overflowX: "auto",
            whiteSpace: "nowrap",
          }}
        >

          {/* LEFT SECTION */}
          <Box
            sx={{
              display: "flex",
              alignItems: "center",
              gap: 1,
              flexShrink: 0,
            }}
          >
            {!isAutoLogin && (
              <Box
                onClick={() => navigate("/")}
                sx={{
                  display: "flex",
                  alignItems: "center",
                  cursor: "pointer",
                }}
              >
                <img
                  src={logo}
                  alt="Pukal Reports"
                  style={{
                    height: 28,
                    objectFit: "contain",
                  }}
                />
              </Box>
            )}

            {showPages &&
              (isAutoLogin ? (
                <Box
                  sx={{
                    minWidth: 160,
                    height: 24,
                    fontSize: "0.75rem",
                    backgroundColor: "#fff",
                    color: "#000000",
                    borderRadius: 0.5,
                    display: "flex",
                    alignItems: "center",
                    px: 1,
                    fontWeight: 600,
                  }}
                >
                  {currentPageLabel}
                </Box>
              ) : (
                <Select
                  size="small"
                  value={selectedPath}
                  onChange={() => { }}
                  sx={{
                    minWidth: isMobile ? 140 : 180,
                    height: 24,
                    fontSize: "0.7rem",
                    backgroundColor: "#fff",
                    borderRadius: 0.5,
                    "& .MuiSelect-select": {
                      py: 0,
                      display: "flex",
                      alignItems: "center",
                    },
                  }}
                >
                  {pages.map((p) => (
                    <MenuItem
                      key={p.path}
                      value={p.path}
                      sx={{ fontSize: "0.7rem" }}
                      onClick={() => navigate(p.path)}
                    >
                      {p.label}
                    </MenuItem>
                  ))}
                </Select>
              ))}

            {templates.length > 0 && (
              <Select
                size="small"
                value={selectedTemplate}
                onChange={(e) => {
                  const reportId = e.target.value;
                  setSelectedTemplate(reportId);

                  if (!reportId) {
                    onReportChange?.(null);
                    return;
                  }
                  const selected = templates.find(
                    (t) => String(t.Report_Id) === String(reportId)
                  );

                  /**
                   * Send selected template data to parent page
                   */
                  if (selected) {
                    onReportChange?.(selected);
                  }
                }}
                displayEmpty
                sx={{
                  minWidth: isMobile ? 130 : 170,
                  height: 24,
                  fontSize: "0.7rem",
                  backgroundColor: "#fff",
                  borderRadius: 0.5,
                  "& .MuiSelect-select": {
                    py: 0,
                    display: "flex",
                    alignItems: "center",
                  },
                }}
              >
                <MenuItem value="">Select Template</MenuItem>

                {templates.map((t) => (
                  <MenuItem key={t.Report_Id} value={t.Report_Id}>
                    {t.Report_Name}
                  </MenuItem>
                ))}
              </Select>
            )}
          </Box>

          {/* DESKTOP COMPANY SWITCH */}
          {!isMobile && user && companies && companies.length > 1 && (
            <Box sx={{ display: "flex", alignItems: "center" }}>
              <Button
                color="inherit"
                endIcon={<ArrowDropDownIcon />}
                onClick={handleCompanyClick}
                sx={{
                  color: "#fff",
                  textTransform: "none",
                  fontWeight: 700,
                }}
              >
                {companyName || "Select Company"}
              </Button>

              <Menu
                anchorEl={anchorElCompany}
                open={openCompanyMenu}
                onClose={() => handleCompanyClose()}
                PaperProps={{ sx: { minWidth: 200 } }}
              >
                {companies.map((company) => {
                  const isSelected = company.id === user?.companyId;

                  return (
                    <MenuItem
                      key={company.id}
                      onClick={() => handleCompanyClose(company)}
                      sx={{
                        fontSize: "0.85rem",
                        fontWeight: isSelected ? 700 : 500,
                        backgroundColor: isSelected
                          ? alpha("#1E3A8A", 0.1)
                          : "inherit",
                      }}
                    >
                      {company.name}
                    </MenuItem>
                  );
                })}
              </Menu>
            </Box>
          )}

          {/* RIGHT SECTION */}
          <Box
            sx={{
              display: "flex",
              alignItems: "center",
              gap: 0.75,
              flexShrink: 0,
            }}
          >
            {infoSlot}

            {onQuickSave && (
              <Tooltip title="Save Template">
                <IconButton
                  size="small"
                  onClick={() => onQuickSave(currentPageLabel)}
                  sx={{
                    height: 24,
                    width: 24,
                    backgroundColor: "#fff",
                    borderRadius: 0.5,
                  }}
                >
                  +
                </IconButton>
              </Tooltip>
            )}

            {toggleMode && onToggleChange && (
              <ToggleButtonGroup
                exclusive
                size="small"
                value={toggleMode}
                onChange={(_, val) => val && onToggleChange(val)}
                sx={{
                  height: 26,
                  backgroundColor: "#f5f5f5",
                  borderRadius: 1,
                  p: 0.25,

                  "& .MuiToggleButton-root": {
                    fontSize: "0.6rem",
                    px: isMobile ? 0.8 : 1.2,
                    py: 0,
                    border: "none",
                    color: "#444",
                    textTransform: "uppercase",
                  },

                  "& .MuiToggleButton-root.Mui-selected": {
                    backgroundColor: "#1e3a8a",
                    color: "#fff",
                    fontWeight: 700,
                    boxShadow: "0 2px 4px rgba(0,0,0,0.25)",
                  },

                  "& .MuiToggleButton-root.Mui-selected:hover": {
                    backgroundColor: "#1e40af",
                  },
                }}
              >
                <ToggleButton value="Abstract">Abstract</ToggleButton>
                <ToggleButton value="Expanded">Expanded</ToggleButton>
              </ToggleButtonGroup>
            )}

            {(onExportPDF || onExportExcel) && (
              <>
                <Tooltip title="Export">
                  <IconButton
                    size="small"
                    onClick={(e) => setAnchorElExport(e.currentTarget)}
                    sx={{
                      height: 24,
                      width: 24,
                      backgroundColor: "#fff",
                      borderRadius: 0.5,
                    }}
                  >
                    <FileDownloadIcon fontSize="small" />
                  </IconButton>
                </Tooltip>

                <Menu
                  anchorEl={anchorElExport}
                  open={openExportMenu}
                  onClose={() => setAnchorElExport(null)}
                >
                  {onExportPDF && (
                    <MenuItem
                      onClick={() => {
                        setAnchorElExport(null);
                        onExportPDF();
                      }}
                    >
                      Export as PDF
                    </MenuItem>
                  )}

                  {onExportExcel && (
                    <MenuItem
                      onClick={() => {
                        setAnchorElExport(null);
                        onExportExcel();
                      }}
                    >
                      Export as Excel
                    </MenuItem>
                  )}
                </Menu>
              </>
            )}

            {settingsSlot}

            {/* MOBILE COMPANY SWITCH */}
            {isMobile && user && companies && companies.length > 1 && (
              <>
                <IconButton
                  size="small"
                  onClick={handleCompanyClick}
                  sx={{ backgroundColor: "#fff" }}
                >
                  <MenuIcon fontSize="small" />
                </IconButton>

                <Menu
                  anchorEl={anchorElCompany}
                  open={openCompanyMenu}
                  onClose={() => handleCompanyClose()}
                  PaperProps={{ sx: { minWidth: 200 } }}
                >
                  {companies.map((company) => {
                    const isSelected = company.id === user?.companyId;

                    return (
                      <MenuItem
                        key={company.id}
                        onClick={() => handleCompanyClose(company)}
                        sx={{
                          fontSize: "0.85rem",
                          fontWeight: isSelected ? 700 : 500,
                        }}
                      >
                        {company.name}
                      </MenuItem>
                    );
                  })}
                </Menu>
              </>
            )}

            {/* LOGOUT */}
            <Tooltip title="Logout">
              <IconButton
                size="small"
                onClick={logout}
                sx={{
                  ml: 0.5,
                  backgroundColor: "#c02222",
                  color: "#fff",
                  width: 28,
                  height: 28,
                  "&:hover": { backgroundColor: "#a61c1c" },
                }}
              >
                <PowerSettingsNewIcon fontSize="small" />
              </IconButton>
            </Tooltip>
          </Box>

        </Toolbar>
      </AppBar>

      <Box sx={{ height: isMobile ? PAGE_HEADER_HEIGHT_MOBILE : PAGE_HEADER_HEIGHT }} />
    </>
  )
};

export default PageHeader;