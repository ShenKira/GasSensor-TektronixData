# Gas Sensor Data Analyzer / 气体传感器数据分析工具

Automated analysis of gas sensor experiment data: segment multi-round experiments, compute rise/fall transition times, and visualize results with interactive plots.

自动化分析气体传感器实验数据：自动切分多轮实验、计算上升/下降过渡时间，并通过交互式图表可视化结果。

---

## Features / 功能

- **Auto segmentation** — Detect experiment rounds from current peaks using Savitzky-Golay smoothing + prominence-based peak detection
- **Edge analysis** — Compute rise/fall times at user-configurable thresholds (default 10%–90%) with linear interpolation for sub-sample precision
- **Dual-axis overlay** — Optionally plot voltage or resistance on a secondary Y-axis with auto-scaling units
- **GUI (PySide6)** — Interactive interface with settings panel, data table, and export (SVG/PNG/Clipboard)
- **i18n** — English / Chinese UI with runtime language switching
- **CLI mode** — Quick command-line analysis with matplotlib output

---

## Project Structure / 项目结构

```
├── gui_main.py              # PySide6 GUI application
├── analyze_core.py          # Core analysis engine
├── analyze_gas_sensor.py    # CLI entry point
├── settings_manager.py      # Settings persistence (settings.json)
├── language_manager.py      # i18n loader
├── language.json            # Translation strings (en/zh)
├── settings.json            # User preferences (auto-generated)
└── testdata.xlsx            # Sample experiment data
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

## License

MIT
