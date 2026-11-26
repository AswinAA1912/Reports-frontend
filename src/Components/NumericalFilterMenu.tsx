import React from "react";
import {
  Menu,
  Box,
  Typography,
  TextField,
  Slider,
  Button,
  ToggleButtonGroup,
  ToggleButton,
  Divider,
} from "@mui/material";

interface NumericalFilterMenuProps {
  anchorEl: HTMLElement | null;
  open: boolean;
  onClose: () => void;
  activeHeader: string | null;
  min: number;
  max: number;
  rangeFilter: Record<string, [number, number]>;
  onRangeChange: (key: string, range: [number, number]) => void;
  onClear: (key: string) => void;
  // Optional aggregation mode (Sum / Avg)
  aggregationMode?: Record<string, "sum" | "avg">;
  onAggregationModeChange?: (key: string, mode: "sum" | "avg") => void;
}

export const NumericalFilterMenu: React.FC<NumericalFilterMenuProps> = ({
  anchorEl,
  open,
  onClose,
  activeHeader,
  min,
  max,
  rangeFilter,
  onRangeChange,
  onClear,
  aggregationMode,
  onAggregationModeChange,
}) => {
  if (!activeHeader) return null;

  const currentRange = rangeFilter[activeHeader] || [min, max];
  const currentMode = aggregationMode ? aggregationMode[activeHeader] || "sum" : "sum";

  const handleSliderChange = (newValue: number[]) => {
    onRangeChange(activeHeader, newValue as [number, number]);
  };

  const handleFromChange = (value: string) => {
    let newFrom = Number(value);
    if (isNaN(newFrom)) return;
    newFrom = Math.max(min, Math.min(newFrom, currentRange[1]));
    handleSliderChange([newFrom, currentRange[1]]);
  };

  const handleToChange = (value: string) => {
    let newTo = Number(value);
    if (isNaN(newTo)) return;
    newTo = Math.min(max, Math.max(newTo, currentRange[0]));
    handleSliderChange([currentRange[0], newTo]);
  };

  return (
    <Menu anchorEl={anchorEl} open={open} onClose={onClose}>
      <Box p={2} sx={{ minWidth: 260 }}>
        {onAggregationModeChange && (
          <Box mb={2}>
            <Typography variant="caption" sx={{ color: "text.secondary", fontWeight: 700, mb: 0.75, display: "block" }}>
              Calculation Type
            </Typography>
            <ToggleButtonGroup
              size="small"
              value={currentMode}
              exclusive
              onChange={(_, newMode) => {
                if (newMode && onAggregationModeChange) {
                  onAggregationModeChange(activeHeader, newMode);
                }
              }}
              fullWidth
              sx={{
                height: 30,
                "& .MuiToggleButton-root": {
                  fontSize: "0.75rem",
                  py: 0.25,
                  fontWeight: 600,
                  textTransform: "none",
                  "&.Mui-selected": {
                    backgroundColor: "#1E3A8A",
                    color: "#fff",
                    "&:hover": {
                      backgroundColor: "#172e6b",
                    },
                  },
                },
              }}
            >
              <ToggleButton value="sum">Sum</ToggleButton>
              <ToggleButton value="avg">Avg</ToggleButton>
            </ToggleButtonGroup>
            <Divider sx={{ my: 1.5 }} />
          </Box>
        )}

        <Typography variant="subtitle2" sx={{ mb: 1.5, fontWeight: 600 }}>
          Filter Range
        </Typography>
        <Box display="flex" alignItems="center" gap={1} mb={2}>
          <TextField
            type="number"
            size="small"
            label="Min"
            value={currentRange[0]}
            onChange={(e) => handleFromChange(e.target.value)}
            inputProps={{ min, max: currentRange[1] }}
            sx={{
              width: 90,
              "& input": {
                py: 0.5,
                fontSize: "0.75rem",
                textAlign: "center",
              },
            }}
          />
          <Typography fontSize={12}>—</Typography>
          <TextField
            type="number"
            size="small"
            label="Max"
            value={currentRange[1]}
            onChange={(e) => handleToChange(e.target.value)}
            inputProps={{ min: currentRange[0], max }}
            sx={{
              width: 90,
              "& input": {
                py: 0.5,
                fontSize: "0.75rem",
                textAlign: "center",
              },
            }}
          />
        </Box>
        <Slider
          value={currentRange}
          min={min}
          max={max}
          step={0.01}
          size="small"
          onChange={(_, newValue) => handleSliderChange(newValue as number[])}
          valueLabelDisplay="auto"
          sx={{
            mb: 1.5,
            "& .MuiSlider-thumb": {
              width: 12,
              height: 12,
            },
          }}
        />
        <Box display="flex" justifyContent="space-between">
          <Button
            size="small"
            onClick={() => {
              onClear(activeHeader);
              onClose();
            }}
            sx={{ textTransform: "none", color: "gray" }}
          >
            Clear
          </Button>
          <Button
            variant="contained"
            size="small"
            onClick={onClose}
            sx={{
              backgroundColor: "#1E3A8A",
              textTransform: "none",
              fontWeight: 600,
            }}
          >
            Close
          </Button>
        </Box>
      </Box>
    </Menu>
  );
};
