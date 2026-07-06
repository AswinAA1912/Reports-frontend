import React, { createContext, useState, useContext } from "react";
import { Outlet } from "react-router-dom";
import { Box } from "@mui/material";

/* ---------------- CONTEXT ---------------- */
interface ToggleContextType {
  toggleMode: "Abstract" | "Expanded";
  setToggleMode: (mode: "Abstract" | "Expanded") => void;
}

const ToggleContext = createContext<ToggleContextType>({
  toggleMode: "Abstract",
  setToggleMode: () => { },
});

export const useToggleMode = () => useContext(ToggleContext);

export interface AppLayoutProps {
  children?: React.ReactNode;
  fullWidth?: boolean;
}

/* ---------------- APP LAYOUT ---------------- */
const AppLayout: React.FC<AppLayoutProps> = ({
  children,
  fullWidth = false,
}) => {
  const [toggleMode, setToggleMode] = useState<"Abstract" | "Expanded">(
    "Abstract"
  );

  return (
    <ToggleContext.Provider value={{ toggleMode, setToggleMode }}>
      {/* ROOT (NO SCROLL HERE) */}
      <Box
        sx={{
          height: "100vh",
          display: "block",
          flexDirection: "column",
          background: fullWidth
            ? "#F1F5F9"
            : "linear-gradient(to bottom, #87CEFA, #B0E0E6)",
        }}
      >
        {/* ✅ SINGLE SCROLL OWNER */}
        <Box
          sx={{
            flex: 1,
            minHeight: 0,
            overflow: "auto",
          }}
        >
          <Box>
            {children ?? <Outlet />}
          </Box>
        </Box>
      </Box>
    </ToggleContext.Provider>
  );
};

export default AppLayout;
