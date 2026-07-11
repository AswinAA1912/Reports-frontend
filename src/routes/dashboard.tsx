import React from "react";
import {
  Box,
  Typography,
  Divider,
  CircularProgress,
  Card,
  CardContent,
  useTheme,
  useMediaQuery
} from "@mui/material";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  Tooltip,
} from "recharts";
import { useNavigate } from "react-router-dom";
import { useAuth } from "../auth/authContext";
import Header from "../Layout/Header";
import { useQuery } from "@tanstack/react-query";
import { MenuService } from "../services/menus.service";
import { DashBoardSalesGraph } from "../services/graphAnalysis.services";

const HEADER_HEIGHT = 64;

const Dashboard: React.FC = () => {
  const { user } = useAuth();
  const navigate = useNavigate();
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down("md"));

  const companyId = user?.companyId ?? undefined;

  const { data: menuList = [], isLoading } = useQuery({
    queryKey: ["dashboard-menus", companyId],
    queryFn: async () => {
      const res = await MenuService.getMenus();

      return res.data.data
        .filter(
          (menu: any) =>
            menu.menu_type === 1 && menu.is_active === 3
        )
        .flatMap((menu: any) =>
          (menu.SubMenu || []).filter(
            (sub: any) => sub.is_active === 3
          )
        )
        .sort(
          (a: any, b: any) =>
            a.display_order - b.display_order
        );
    },
    enabled: !!companyId,
  });

  const today = new Date();

  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);

  const formatDate = (date: Date) =>
    date.toISOString().split("T")[0];

  const fromDate = formatDate(firstDay);
  const toDate = formatDate(today);

  const { data: graphData, isLoading: graphLoading } = useQuery({
    queryKey: ["dashboard-graph", user?.companyId, fromDate, toDate],
    queryFn: async () => {
      const res = await DashBoardSalesGraph.getDashboardGraph({
        Fromdate: fromDate,
        Todate: toDate,
        Company_Id: companyId,
      });
      return res.data.data;
    },
    enabled: !!user?.companyId,
  });

  const latestDate = graphData?.DayWise?.length
    ? new Date(
      Math.max(
        ...graphData.DayWise.map((d: any) =>
          new Date(d.Invoice_Date).getTime()
        )
      )
    )
    : new Date();

  const currentMonth = latestDate.getMonth() + 1;
  const currentYear = latestDate.getFullYear();
  const currentDay = latestDate.getDate();

  const monthData =
    graphData?.DayWise?.filter((d: any) => {
      const dateStr = d.Invoice_Date.split("T")[0];
      const [year, month, day] = dateStr.split("-").map(Number);

      return (
        year === currentYear &&
        month === currentMonth &&
        day <= currentDay
      );
    }) || [];

  const totalInvoices = monthData.reduce(
    (sum: number, item: any) => sum + item.Invoice_Count,
    0
  );

  const totalSales = monthData.reduce(
    (sum: number, item: any) => sum + item.Total_Invoice_value,
    0
  );

  const chartData = monthData.map((d: any) => ({
    day: Number(d.Invoice_Date.split("T")[0].split("-")[2]),
    value: d.Total_Invoice_value,
  }));

  return (
    <>
      <Header headerColor="#1E3A8A" showSearch={false} />

      <Box
        sx={{
          display: "flex",
          flexDirection: isMobile ? "column" : "row",
          height: isMobile ? "auto" : `calc(115vh - ${HEADER_HEIGHT}px)`,
          width: "100%",
          overflow: "hidden",
          backgroundColor: "#cfe6ec",
        }}
      >
        {/* LEFT PANEL */}
        <Box
          sx={{
            width: isMobile ? "100%" : 360,
            px: isMobile ? 2 : 5,
            py: isMobile ? 3 : 0,
            display: "flex",
            flexDirection: "column",
            justifyContent: "center",
            gap: 3
          }}
        >
          <Box>
            <Typography
              sx={{
                fontSize: isMobile ? 20 : 32,
                fontWeight: 400,
                mb: 1,
                whiteSpace: isMobile ? "normal" : "nowrap",
                overflow: "hidden",
                textOverflow: "ellipsis",
                wordBreak: "break-word",
              }}
            >
              Welcome {user?.Name || "admin"},
            </Typography>

            <Typography
              sx={{
                fontSize: 22,
                fontWeight: 700,
                letterSpacing: 1,
                fontFamily: "serif",
              }}
            >
              {user?.Company_Name || "PUKAL FOODS PVT LTD"}
            </Typography>
          </Box>

          {/* SALES ERP CARD */}
          <Card
            onClick={() => navigate("/graphanalysis")}
            sx={{
              cursor: "pointer",
              borderRadius: 3,
              background: "linear-gradient(135deg,#1E3A8A,#2563EB)",
              color: "#fff",
              boxShadow: "0 8px 20px rgba(0,0,0,0.25)",
              transition: "all 0.3s ease",
              "&:hover": {
                transform: "translateY(-6px)",
                boxShadow: "0 12px 28px rgba(0,0,0,0.35)",
              },
            }}
          >
            <CardContent
              sx={{
                display: "flex",
                flexDirection: isMobile ? "column" : "row",
                justifyContent: "space-between",
                alignItems: isMobile ? "flex-start" : "center",
                gap: isMobile ? 2 : 0,
              }}
            >
              {/* LEFT SIDE DATA */}
              <Box sx={{ minWidth: 0 }}>
                <Typography
                  sx={{
                    fontSize: 14,
                    opacity: 0.9,
                    mb: 1,
                    letterSpacing: 1,
                  }}
                >
                  TOTAL SALES
                </Typography>

                {graphLoading ? (
                  <CircularProgress size={22} sx={{ color: "#fff" }} />
                ) : (
                  <>
                    <Typography
                      sx={{
                        fontSize: 28,
                        fontWeight: 700,
                      }}
                    >
                      ₹ {new Intl.NumberFormat("en-IN").format(totalSales)}
                    </Typography>

                    <Typography
                      sx={{
                        fontSize: 14,
                        opacity: 0.85,
                        mt: 1,
                      }}
                    >
                      {totalInvoices} Invoices
                    </Typography>
                  </>
                )}

                <Typography sx={{ fontSize: 12, opacity: 0.8, mt: 1 }}>
                  This Month
                </Typography>
              </Box>

              {/* RIGHT SIDE MINI GRAPH */}
              <Box sx={{
                width: isMobile ? "100%" : 120,
                height: isMobile ? 100 : 70,
              }}>
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={chartData}>

                    <defs>
                      <linearGradient id="salesLine" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#ffffff" stopOpacity={1} />
                        <stop offset="100%" stopColor="#ffffff" stopOpacity={0.2} />
                      </linearGradient>
                    </defs>

                    <XAxis dataKey="day" hide />

                    <Tooltip
                      contentStyle={{
                        background: "#1E3A8A",
                        border: "none",
                        borderRadius: "6px",
                        color: "#fff"
                      }}
                      formatter={(value: any) =>
                        `₹ ${new Intl.NumberFormat("en-IN").format(Number(value || 0))}`
                      }
                    />

                    <Line
                      type="monotone"
                      dataKey="value"
                      stroke="#ffffff"
                      strokeWidth={3}
                      dot={false}
                      activeDot={{ r: 5 }}
                      isAnimationActive={true}
                      animationDuration={1200}
                    />

                  </LineChart>
                </ResponsiveContainer>
              </Box>
            </CardContent>
          </Card>
        </Box>

        {!isMobile && <Divider orientation="vertical" flexItem />}

        {/* RIGHT MENU */}
        <Box
          sx={{
            flex: 1,
            px: isMobile ? 2 : 6,
            py: isMobile ? 4 : 4,
            display: "flex",
            flexDirection: "column",
            justifyContent: "flex-start",
            overflowY: "auto",
            maxHeight: "100%",
          }}
        >
          <Typography fontSize={22} fontWeight={700} mb={2}>
            MENU
          </Typography>

          <Divider sx={{ mb: 3 }} />

          {isLoading ? (
            <Box textAlign="center">
              <CircularProgress />
            </Box>
          ) : (
            <Box
              sx={{
                display: "grid",
                gridTemplateColumns: isMobile
                  ? "1fr"
                  : "repeat(auto-fill, minmax(220px, 1fr))",
                gap: 2,
              }}
            >
              {menuList.map((item: any) => (
                <Box
                  key={item.id}
                  onClick={() => navigate(item.rUrl)}
                  sx={{
                    cursor: "pointer",
                    p: 2,
                    borderRadius: 2,
                    boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
                    transition: "all 0.2s ease",
                    "&:hover": {
                      transform: "translateY(-3px)",
                      boxShadow: "0 6px 16px rgba(0,0,0,0.15)",
                      background: "#b1c6da",
                    },
                  }}
                >
                  <Typography fontSize={16} fontWeight={600}>
                    {item.name}
                  </Typography>
                </Box>
              ))}
            </Box>
          )}
        </Box>
      </Box>
    </>
  );
};

export default Dashboard;