# Gas Sensor Data Analyzer / 气体传感器数据分析工具

Automated analysis of gas sensor experiment data: segment multi-round experiments, detect peak types (overshoot vs. saturation), compute rise/fall transition times, and visualize results with interactive plots.

自动化分析气体传感器实验数据：自动切分多轮实验、识别峰值类型（过冲峰/饱和峰）、计算上升/下降过渡时间，并通过交互式图表可视化结果。

---

## Features / 功能

- **Auto segmentation** — Detect experiment rounds from current peaks using Savitzky-Golay smoothing + prominence-based peak detection
- **Peak type detection** — Automatically distinguish **overshoot peaks** (过冲峰) from **saturation peaks** (饱和峰) using two-phase descent analysis
  - Overshoot: current immediately descends after peak, then drops rapidly when gas stops
  - Saturation: current plateaus near peak, drops rapidly only when gas stops
- **Overshoot descent analysis** — For overshoot peaks: compute I_90_Stable (steady-state current), slow descent time, and overshoot descent time
- **Edge analysis** — Compute rise/fall times at user-configurable thresholds (default 10%–90%) with linear interpolation for sub-sample precision
  - For overshoot peaks, fall edge is computed on the rapid drop phase only (not the full descent)
- **Dual-axis overlay** — Optionally plot voltage or resistance on a secondary Y-axis with auto-scaling units
- **GUI (PySide6)** — Interactive interface with settings panel, data table, peak type indicator, and export (SVG/PNG/Clipboard)
- **Batch processing** — Process multiple files with peak type annotation in Excel export ("-过冲峰" / "-饱和峰")
- **i18n** — English / Chinese UI with runtime language switching
- **CLI mode** — Quick command-line analysis with matplotlib output

---

## Project Structure / 项目结构

```
├── gui_main.py              # PySide6 GUI application
├── analyze_core.py          # Core analysis engine (peak detection + edge analysis)
├── analyze_gas_sensor.py    # CLI entry point
├── settings_manager.py      # Settings persistence (settings.json)
├── language_manager.py      # i18n loader
├── language.json            # Translation strings (en/zh)
├── settings.json            # User preferences (auto-generated)
├── testdata_peak.xlsx       # Sample overshoot peak data (11 rounds)
├── testdata_saturation.xlsx # Sample saturation peak data (13 rounds)
└── Vibe_Coding_Outline.md   # Detailed module documentation
```

## Quick Start / 快速开始

```bash
# GUI mode
python gui_main.py

# CLI mode
python analyze_gas_sensor.py
```

## Requirements / 依赖

```
numpy
scipy
openpyxl
matplotlib
PySide6
```

## Data Format / 数据格式

The input `.xlsx` file should contain:

| Column | Content |
|--------|---------|
| 1 | Index |
| 2 | Time (s) |
| 3 | Voltage (V) |
| 4 | Current (A) |
| 5 | Resistance (Ω) |

## Data Table Columns / 数据表列

| Column | Content |
|--------|---------|
| Round | Experiment round number |
| Peak I | Maximum smoothed current (auto-scaled unit) |
| Idle I | Baseline current (trimmed mean) |
| Response | Peak I / Idle I ratio |
| Rise Time (s) | 10%→90% rise transition time |
| Fall Time (s) | 90%→10% fall transition time (rapid drop for overshoot) |
| Peak Type | "过冲峰" (Overshoot) or "饱和峰" (Saturation) |
| Slow Descent Time (s) | Peak to I_90_Stable crossing time (overshoot only) |
| I_90_Stable | Steady-state current after overshoot slow descent |
| Response(Stable) | I_90_Stable / Idle I ratio |
| I_upper% / I_lower% | Current at rise/fall edge crossing points |

## License

MIT
