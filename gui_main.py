"""
气体传感器实验数据分析 - 图形界面版本。
使用 PySide6 构建 GUI，集成 Matplotlib 绘图。
"""

import sys
import io
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QFileDialog,
    QDialog, QFormLayout, QSpinBox, QDoubleSpinBox, QCheckBox,
    QDialogButtonBox, QMessageBox, QHeaderView, QLabel, QGroupBox,
    QSplitter, QSizePolicy, QFrame, QComboBox
)
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QAction, QIcon, QClipboard, QFont

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.lines import Line2D

from settings_manager import SettingsManager, AppSettings, AnalysisSettings
from analyze_core import run_analysis, AnalysisResult
from language_manager import LanguageManager


# Windows console UTF-8 output
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


# 蓝白配色扁平化样式表
FLAT_STYLE = """
QMainWindow, QDialog {
    background-color: #f5f7fa;
}
QWidget {
    font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
    font-size: 11px;
}
QGroupBox {
    background-color: white;
    border: 1px solid #d0d7de;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 15px;
    font-weight: bold;
    color: #24292f;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}
QPushButton {
    background-color: #0969da;
    color: white;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
    min-width: 80px;
}
QPushButton:hover {
    background-color: #0860ca;
}
QPushButton:pressed {
    background-color: #0757b5;
}
QPushButton:disabled {
    background-color: #8c959f;
}
QTableWidget {
    background-color: white;
    border: 1px solid #d0d7de;
    border-radius: 6px;
    gridline-color: #e1e4e8;
    selection-background-color: #ddf4ff;
    selection-color: #24292f;
}
QTableWidget::item {
    padding: 5px;
}
QHeaderView::section {
    background-color: #f6f8fa;
    border: none;
    border-bottom: 2px solid #0969da;
    padding: 8px;
    font-weight: bold;
    color: #24292f;
}
QCheckBox {
    spacing: 8px;
    color: #24292f;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #d0d7de;
    border-radius: 4px;
    background-color: white;
}
QCheckBox::indicator:checked {
    background-color: #0969da;
    border-color: #0969da;
}
QCheckBox::indicator:disabled {
    background-color: #f6f8fa;
    border-color: #e1e4e8;
}
QRadioButton {
    spacing: 8px;
    color: #24292f;
}
QRadioButton::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #d0d7de;
    border-radius: 9px;
    background-color: white;
}
QRadioButton::indicator:checked {
    background-color: #0969da;
    border-color: #0969da;
}
QComboBox {
    border: 1px solid #d0d7de;
    border-radius: 6px;
    padding: 5px 10px;
    background-color: white;
    color: #24292f;
}
QComboBox:hover {
    border-color: #0969da;
}
QComboBox::drop-down {
    border: none;
}
QLabel {
    color: #24292f;
}
QMenuBar {
    background-color: white;
    border-bottom: 1px solid #d0d7de;
}
QMenuBar::item:selected {
    background-color: #ddf4ff;
}
QMenu {
    background-color: white;
    border: 1px solid #d0d7de;
}
QMenu::item:selected {
    background-color: #ddf4ff;
}
QStatusBar {
    background-color: #f6f8fa;
    border-top: 1px solid #d0d7de;
}
QSplitter::handle {
    background-color: #d0d7de;
}
QSpinBox, QDoubleSpinBox {
    border: 1px solid #d0d7de;
    border-radius: 4px;
    padding: 4px;
    background-color: white;
}
QSpinBox:hover, QDoubleSpinBox:hover {
    border-color: #0969da;
}
QSpinBox::up-button, QDoubleSpinBox::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    border-left: 1px solid #d0d7de;
    border-bottom: 1px solid #d0d7de;
    border-top-right-radius: 4px;
    background-color: #f6f8fa;
    width: 20px;
}
QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover {
    background-color: #ddf4ff;
}
QSpinBox::down-button, QDoubleSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    border-left: 1px solid #d0d7de;
    border-top: 1px solid #d0d7de;
    border-bottom-right-radius: 4px;
    background-color: #f6f8fa;
    width: 20px;
}
QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {
    background-color: #ddf4ff;
}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-bottom: 7px solid #24292f;
}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 7px solid #24292f;
}
"""


def get_current_unit_and_scale(max_val: float) -> Tuple[str, float]:
    """
    根据电流最大值自动选择合适的单位和缩放因子。
    返回 (单位字符串, 缩放因子)。
    规则：当最大值 > 5000X 时，升级到下一个单位。
    """
    abs_max = abs(max_val)
    if abs_max == 0:
        return 'A', 1.0

    # 从 nA 开始判断
    # 1 A = 1e6 uA = 1e9 nA
    if abs_max < 5e-6:  # < 5 uA，使用 nA
        return 'nA', 1e9
    elif abs_max < 5e-3:  # < 5 mA，使用 uA
        return 'μA', 1e6
    elif abs_max < 5.0:  # < 5 A，使用 mA
        return 'mA', 1e3
    else:
        return 'A', 1.0


def get_voltage_unit_and_scale(max_val: float) -> Tuple[str, float]:
    """
    根据电压最大值自动选择合适的单位和缩放因子。
    返回 (单位字符串, 缩放因子)。
    规则：当最大值 > 5000X 时，升级到下一个单位。
    """
    abs_max = abs(max_val)
    if abs_max == 0:
        return 'V', 1.0

    # 从 mV 开始判断
    if abs_max < 5e-3:  # < 5 mV，使用 mV
        return 'mV', 1e3
    elif abs_max < 5.0:  # < 5 V，使用 V
        return 'V', 1.0
    elif abs_max < 5e3:  # < 5 kV，使用 kV
        return 'kV', 1e-3
    else:
        return 'MV', 1e-6


def get_resistance_unit_and_scale(max_val: float) -> Tuple[str, float]:
    """
    根据电阻最大值自动选择合适的单位和缩放因子。
    返回 (单位字符串, 缩放因子)。
    规则：当最大值 > 5000X 时，升级到下一个单位。
    """
    abs_max = abs(max_val)
    if abs_max == 0:
        return 'Ω', 1.0

    # 从 Ω 开始判断
    if abs_max < 5e3:  # < 5 kΩ，使用 Ω
        return 'Ω', 1.0
    elif abs_max < 5e6:  # < 5 MΩ，使用 kΩ
        return 'kΩ', 1e-3
    elif abs_max < 5e9:  # < 5 GΩ，使用 MΩ
        return 'MΩ', 1e-6
    else:
        return 'GΩ', 1e-9


def format_current(val: float, scale: float) -> str:
    """格式化电流值，使用指定的缩放因子。"""
    return f"{val * scale:.4f}"


class SettingsDialog(QDialog):
    """设置对话框"""

    def __init__(self, settings: AppSettings, lang: LanguageManager, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.lang = lang
        self.setWindowTitle(self.lang.get("dialog_settings"))
        self.setMinimumWidth(400)
        self._setup_ui()

    def _tr(self, key: str, **kwargs) -> str:
        return self.lang.get(key, **kwargs)

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # 分析参数组
        analysis_group = QGroupBox(self._tr("group_analysis"))
        analysis_layout = QFormLayout()

        self.rise_lower = QDoubleSpinBox()
        self.rise_lower.setRange(0.1, 99.9)
        self.rise_lower.setValue(self.settings.analysis.rise_lower_percent)
        self.rise_lower.setSuffix(" %")
        analysis_layout.addRow(self._tr("label_rise_lower"), self.rise_lower)

        self.rise_upper = QDoubleSpinBox()
        self.rise_upper.setRange(0.1, 99.9)
        self.rise_upper.setValue(self.settings.analysis.rise_upper_percent)
        self.rise_upper.setSuffix(" %")
        analysis_layout.addRow(self._tr("label_rise_upper"), self.rise_upper)

        self.fall_upper = QDoubleSpinBox()
        self.fall_upper.setRange(0.1, 99.9)
        self.fall_upper.setValue(self.settings.analysis.fall_upper_percent)
        self.fall_upper.setSuffix(" %")
        analysis_layout.addRow(self._tr("label_fall_upper"), self.fall_upper)

        self.fall_lower = QDoubleSpinBox()
        self.fall_lower.setRange(0.1, 99.9)
        self.fall_lower.setValue(self.settings.analysis.fall_lower_percent)
        self.fall_lower.setSuffix(" %")
        analysis_layout.addRow(self._tr("label_fall_lower"), self.fall_lower)

        analysis_group.setLayout(analysis_layout)
        layout.addWidget(analysis_group)

        # 峰值检测参数组
        detection_group = QGroupBox(self._tr("group_detection"))
        detection_layout = QFormLayout()

        self.prominence = QDoubleSpinBox()
        self.prominence.setRange(0.01, 5.0)
        self.prominence.setValue(self.settings.analysis.prominence)
        self.prominence.setSingleStep(0.05)
        self.prominence.setDecimals(2)
        detection_layout.addRow(self._tr("label_prominence"), self.prominence)

        self.distance = QSpinBox()
        self.distance.setRange(10, 1000)
        self.distance.setValue(self.settings.analysis.distance)
        self.distance.setSuffix(" pts")
        detection_layout.addRow(self._tr("label_distance"), self.distance)

        self.baseline_fraction = QDoubleSpinBox()
        self.baseline_fraction.setRange(0.001, 0.5)
        self.baseline_fraction.setValue(self.settings.analysis.baseline_fraction)
        self.baseline_fraction.setSingleStep(0.005)
        self.baseline_fraction.setDecimals(3)
        detection_layout.addRow(self._tr("label_baseline"), self.baseline_fraction)

        self.peak_threshold = QDoubleSpinBox()
        self.peak_threshold.setRange(1e-12, 1e-3)
        self.peak_threshold.setValue(self.settings.analysis.peak_threshold)
        self.peak_threshold.setSingleStep(1e-8)
        self.peak_threshold.setDecimals(10)
        self.peak_threshold.setSuffix(" A")
        detection_layout.addRow(self._tr("label_peak_thresh"), self.peak_threshold)

        # 重置按钮
        reset_btn = QPushButton(self._tr("btn_reset"))
        reset_btn.setStyleSheet("""
            QPushButton {
                background-color: #6e7781;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #57606a;
            }
        """)
        reset_btn.clicked.connect(self._reset_detection)
        detection_layout.addRow("", reset_btn)

        detection_group.setLayout(detection_layout)
        layout.addWidget(detection_group)

        # 绘图参数组
        plot_group = QGroupBox(self._tr("group_plot"))
        plot_layout = QFormLayout()

        self.title_size = QSpinBox()
        self.title_size.setRange(8, 36)
        self.title_size.setValue(self.settings.plot.font_size_title)
        plot_layout.addRow(self._tr("label_title_size"), self.title_size)

        self.title_bold = QCheckBox()
        self.title_bold.setChecked(self.settings.plot.title_bold)
        plot_layout.addRow(self._tr("label_title_bold"), self.title_bold)

        self.axis_label_size = QSpinBox()
        self.axis_label_size.setRange(8, 24)
        self.axis_label_size.setValue(self.settings.plot.font_size_axis_label)
        plot_layout.addRow(self._tr("label_axis_size"), self.axis_label_size)

        self.axis_label_bold = QCheckBox()
        self.axis_label_bold.setChecked(self.settings.plot.axis_label_bold)
        plot_layout.addRow(self._tr("label_axis_bold"), self.axis_label_bold)

        self.tick_size = QSpinBox()
        self.tick_size.setRange(6, 18)
        self.tick_size.setValue(self.settings.plot.font_size_tick)
        plot_layout.addRow(self._tr("label_tick_size"), self.tick_size)

        self.legend_size = QSpinBox()
        self.legend_size.setRange(6, 18)
        self.legend_size.setValue(self.settings.plot.font_size_legend)
        plot_layout.addRow(self._tr("label_legend_size"), self.legend_size)

        self.annotation_size = QSpinBox()
        self.annotation_size.setRange(6, 16)
        self.annotation_size.setValue(self.settings.plot.font_size_annotation)
        plot_layout.addRow(self._tr("label_annot_size"), self.annotation_size)

        plot_group.setLayout(plot_layout)
        layout.addWidget(plot_group)

        # 界面缩放参数组
        scale_group = QGroupBox(self._tr("group_gui_scale"))
        scale_layout = QFormLayout()

        self.gui_scale = QDoubleSpinBox()
        self.gui_scale.setRange(0.5, 3.0)
        self.gui_scale.setValue(self.settings.gui_scale)
        self.gui_scale.setSingleStep(0.1)
        self.gui_scale.setSuffix(" x")
        scale_layout.addRow(self._tr("label_gui_scale"), self.gui_scale)

        scale_group.setLayout(scale_layout)
        layout.addWidget(scale_group)

        # 语言切换参数组
        lang_group = QGroupBox(self._tr("group_language"))
        lang_layout = QFormLayout()

        self.combo_language = QComboBox()
        for code in self.lang.get_available_languages():
            self.combo_language.addItem(self.lang.get_language_name(code), code)
        current_idx = self.lang.get_available_languages().index(self.lang.language)
        self.combo_language.setCurrentIndex(current_idx)
        lang_layout.addRow(self._tr("label_language"), self.combo_language)

        lang_group.setLayout(lang_layout)
        layout.addWidget(lang_group)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _reset_detection(self):
        """重置峰值检测参数为默认值"""
        defaults = AnalysisSettings()
        self.prominence.setValue(defaults.prominence)
        self.distance.setValue(defaults.distance)
        self.baseline_fraction.setValue(defaults.baseline_fraction)
        self.peak_threshold.setValue(defaults.peak_threshold)

    def get_settings(self) -> dict:
        """返回更新后的设置"""
        return {
            'analysis': {
                'rise_lower_percent': self.rise_lower.value(),
                'rise_upper_percent': self.rise_upper.value(),
                'fall_upper_percent': self.fall_upper.value(),
                'fall_lower_percent': self.fall_lower.value(),
                'prominence': self.prominence.value(),
                'distance': self.distance.value(),
                'baseline_fraction': self.baseline_fraction.value(),
                'peak_threshold': self.peak_threshold.value(),
            },
            'plot': {
                'font_size_title': self.title_size.value(),
                'title_bold': self.title_bold.isChecked(),
                'font_size_axis_label': self.axis_label_size.value(),
                'axis_label_bold': self.axis_label_bold.isChecked(),
                'font_size_tick': self.tick_size.value(),
                'font_size_legend': self.legend_size.value(),
                'font_size_annotation': self.annotation_size.value(),
                'line_width_raw': self.settings.plot.line_width_raw,
                'line_width_smooth': self.settings.plot.line_width_smooth,
                'line_width_edge': self.settings.plot.line_width_edge,
            },
            'gui_scale': self.gui_scale.value(),
            'language': self.combo_language.currentData(),
        }

    def get_analysis_changed(self) -> bool:
        """检查分析参数是否改变"""
        return (
            self.rise_lower.value() != self.settings.analysis.rise_lower_percent or
            self.rise_upper.value() != self.settings.analysis.rise_upper_percent or
            self.fall_upper.value() != self.settings.analysis.fall_upper_percent or
            self.fall_lower.value() != self.settings.analysis.fall_lower_percent or
            self.prominence.value() != self.settings.analysis.prominence or
            self.distance.value() != self.settings.analysis.distance or
            self.baseline_fraction.value() != self.settings.analysis.baseline_fraction or
            self.peak_threshold.value() != self.settings.analysis.peak_threshold
        )


class PlotCanvas(FigureCanvas):
    """Matplotlib 绘图画布"""

    def __init__(self, parent=None):
        self.fig = Figure(figsize=(12, 6), dpi=100)
        super().__init__(self.fig)
        self.setParent(parent)
        self.ax = self.fig.add_subplot(111)
        self.ax_right = None  # 右侧Y轴，初始为None
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def plot_data(self, result: AnalysisResult, settings: AppSettings,
                  show_voltage: bool = False, show_resistance: bool = False,
                  left_scale: str = 'linear', right_scale: str = 'linear'):
        """绘制分析结果"""
        self.fig.clear()

        times = result.times
        currents = result.currents
        voltages = result.voltages
        resistances = result.resistances
        currents_smooth = result.currents_smooth
        rounds = result.rounds
        plot_settings = settings.plot

        # 自动选择电流单位
        current_unit, current_scale = get_current_unit_and_scale(np.max(currents))

        # 设置字体
        font_family = 'Times New Roman'

        # 转换电流到合适的单位
        currents_scaled = currents * current_scale
        currents_smooth_scaled = currents_smooth * current_scale

        # 创建左侧Y轴
        self.ax = self.fig.add_subplot(111)

        # 绘制原始数据和平滑数据
        self.ax.plot(
            times, currents_scaled,
            'k-', linewidth=plot_settings.line_width_raw, alpha=0.4,
            label='Raw data'
        )
        self.ax.plot(
            times, currents_smooth_scaled,
            'b-', linewidth=plot_settings.line_width_smooth, alpha=0.7,
            label='Smoothed'
        )

        # 绘制上升沿和下降沿
        colors_rise = matplotlib.cm.Greens(np.linspace(0.4, 0.9, len(rounds)))
        colors_fall = matplotlib.cm.Reds(np.linspace(0.4, 0.9, len(rounds)))

        for i, r in enumerate(rounds):
            start, peak, end = r.start_idx, r.peak_idx, r.end_idx

            # 上升沿（绿色）—— 使用平滑数据，保证与峰值检测/分析结果的一致性
            self.ax.plot(
                times[start:peak + 1], currents_smooth_scaled[start:peak + 1],
                color=colors_rise[i], linewidth=plot_settings.line_width_edge,
            )
            # 下降沿（红色）
            self.ax.plot(
                times[peak:end + 1], currents_smooth_scaled[peak:end + 1],
                color=colors_fall[i], linewidth=plot_settings.line_width_edge,
            )

            # 峰值标记 —— 使用平滑数据，与 find_round_boundaries 的峰值检测一致
            self.ax.plot(times[peak], currents_smooth_scaled[peak], '*', color='red', markersize=10)

            # 轮次标注
            self.ax.annotate(
                f'#{r.round_num}',
                xy=(times[peak], currents_smooth_scaled[peak]),
                xytext=(5, 10), textcoords='offset points',
                fontsize=plot_settings.font_size_annotation,
                color='darkred', fontfamily=font_family,
            )

            # 标记百分比交叉点（使用平滑数据值，与插值计算一致）
            if r.rise:
                for key, marker in [('t_lower', 'o'), ('t_upper', 's')]:
                    t_val = getattr(r.rise, key)
                    idx = np.argmin(np.abs(times - t_val))
                    self.ax.plot(t_val, currents_smooth_scaled[idx], marker, color='green', markersize=5)
            if r.fall:
                for key, marker in [('t_upper', 'o'), ('t_lower', 's')]:
                    t_val = getattr(r.fall, key)
                    idx = np.argmin(np.abs(times - t_val))
                    self.ax.plot(t_val, currents_smooth_scaled[idx], marker, color='orange', markersize=5)

        # 设置左侧Y轴
        self.ax.set_ylabel(
            f'Current ({current_unit})',
            fontsize=plot_settings.font_size_axis_label,
            fontweight='bold' if plot_settings.axis_label_bold else 'normal',
            fontfamily=font_family,
        )

        # 设置左侧Y轴缩放
        self.ax.set_yscale(left_scale)

        # 创建右侧Y轴（如果需要）
        self.ax_right = None
        if show_voltage or show_resistance:
            self.ax_right = self.ax.twinx()
            # 将右侧Y轴置于左侧Y轴之后，防止其线条遮挡图例
            self.ax_right.set_zorder(self.ax.get_zorder() - 1)
            self.ax.patch.set_visible(False)

            if show_voltage:
                # 自动选择电压单位
                voltage_unit, voltage_scale = get_voltage_unit_and_scale(np.max(voltages))
                voltages_scaled = voltages * voltage_scale

                # 计算 RMS 和 RMSE（标准差）
                voltage_rms = np.sqrt(np.mean(voltages**2))
                voltage_rmse = np.std(voltages)
                voltage_rms_scaled = voltage_rms * voltage_scale
                voltage_rmse_scaled = voltage_rmse * voltage_scale

                self.ax_right.plot(
                    times, voltages_scaled,
                    'r-', linewidth=1.5, alpha=0.7, label='Voltage'
                )
                # 设置电压Y轴从0开始，线性坐标
                self.ax_right.set_ylim(bottom=0)
                self.ax_right.set_ylabel(
                    f'Voltage ({voltage_unit})',
                    fontsize=plot_settings.font_size_axis_label,
                    fontweight='bold' if plot_settings.axis_label_bold else 'normal',
                    fontfamily=font_family,
                    color='#c92a2a',
                )
                self.ax_right.tick_params(axis='y', labelcolor='#c92a2a')
                self.ax_right.set_yscale('linear')

                # 当 RMSE < 0.25 * RMS 时，限制Y轴上限为最大值*1.3
                if voltage_rmse < 0.25 * voltage_rms:
                    self.ax_right.set_ylim(top=np.max(voltages_scaled) * 1.3)

                # 在图上部偏左显示 RMS 和 RMSE
                self.ax_right.text(
                    0.02, 0.97,
                    f'RMS = {voltage_rms_scaled:.2f} {voltage_unit}\n'
                    f'RMSE = {voltage_rmse_scaled:.2f} {voltage_unit}',
                    transform=self.ax_right.transAxes,
                    fontsize=plot_settings.font_size_legend,
                    fontfamily=font_family,
                    verticalalignment='top',
                    color='#c92a2a',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                              edgecolor='#c92a2a', alpha=0.8),
                )

            elif show_resistance:
                # 自动选择电阻单位
                resistance_unit, resistance_scale = get_resistance_unit_and_scale(np.max(resistances))
                resistances_scaled = resistances * resistance_scale

                self.ax_right.plot(
                    times, resistances_scaled,
                    'm-', linewidth=1.5, alpha=0.7, label='Resistance'
                )
                self.ax_right.set_ylabel(
                    f'Resistance ({resistance_unit})',
                    fontsize=plot_settings.font_size_axis_label,
                    fontweight='bold' if plot_settings.axis_label_bold else 'normal',
                    fontfamily=font_family,
                    color='#862e9c',
                )
                self.ax_right.tick_params(axis='y', labelcolor='#862e9c')

            # 设置右侧Y轴缩放
            self.ax_right.set_yscale(right_scale)

        # 图例
        legend_elements = [
            Line2D([0], [0], color='green', linewidth=2, label='Rising edge'),
            Line2D([0], [0], color='red', linewidth=2, label='Falling edge'),
            Line2D([0], [0], marker='*', color='red', linestyle='None', markersize=10, label='Peak'),
            Line2D([0], [0], marker='o', color='green', linestyle='None', markersize=5,
                   label=f'{settings.analysis.rise_lower_percent:.0f}% point (rise)'),
            Line2D([0], [0], marker='s', color='green', linestyle='None', markersize=5,
                   label=f'{settings.analysis.rise_upper_percent:.0f}% point (rise)'),
        ]

        # 添加右侧Y轴的图例
        if show_voltage:
            legend_elements.append(
                Line2D([0], [0], color='#c92a2a', linewidth=1.5, label='Voltage')
            )
        elif show_resistance:
            legend_elements.append(
                Line2D([0], [0], color='#862e9c', linewidth=1.5, label='Resistance')
            )

        legend = self.ax.legend(
            handles=legend_elements, loc='upper right',
            fontsize=plot_settings.font_size_legend, prop={'family': font_family},
        )
        legend.set_zorder(10)

        # 设置轴标签
        self.ax.set_xlabel(
            'Time (s)',
            fontsize=plot_settings.font_size_axis_label,
            fontweight='bold' if plot_settings.axis_label_bold else 'normal',
            fontfamily=font_family,
        )
        self.ax.set_title(
            'Gas Sensor Experiment',
            fontsize=plot_settings.font_size_title,
            fontweight='bold' if plot_settings.title_bold else 'normal',
            fontfamily=font_family,
        )

        # 设置刻度字体
        self.ax.tick_params(axis='both', labelsize=plot_settings.font_size_tick)
        for label in self.ax.get_xticklabels() + self.ax.get_yticklabels():
            label.set_fontfamily('Times New Roman')

        if self.ax_right:
            self.ax_right.tick_params(axis='y', labelsize=plot_settings.font_size_tick)
            for label in self.ax_right.get_yticklabels():
                label.set_fontfamily('Times New Roman')

        self.ax.grid(True, which='both', alpha=0.3)
        self.fig.tight_layout()
        self.draw()

    def get_figure(self) -> Figure:
        """返回 Figure 对象"""
        return self.fig


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.settings_manager = SettingsManager()
        self.lang = LanguageManager(language=self.settings_manager.get_settings().language)
        self.analysis_result: Optional[AnalysisResult] = None
        self.current_file_path: Optional[str] = None
        self._setup_ui()
        self._setup_menu()
        self._setup_status_bar()

    def _tr(self, key: str, **kwargs) -> str:
        return self.lang.get(key, **kwargs)

    def _setup_ui(self):
        """设置界面布局"""
        self.setWindowTitle(self._tr("window_title"))
        self.setMinimumSize(1400, 900)

        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局（水平）
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(10)

        # 左侧控制面板
        left_panel = QFrame()
        left_panel.setFixedWidth(200)
        left_panel.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #d0d7de;
                border-radius: 6px;
            }
        """)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 15, 10, 10)

        # 数据叠加选项组
        self.overlay_group = QGroupBox(self._tr("group_overlay"))
        overlay_layout = QVBoxLayout(self.overlay_group)

        self.chk_voltage = QCheckBox(self._tr("chk_voltage"))
        self.chk_voltage.stateChanged.connect(self._on_overlay_changed)
        overlay_layout.addWidget(self.chk_voltage)

        self.chk_resistance = QCheckBox(self._tr("chk_resistance"))
        self.chk_resistance.stateChanged.connect(self._on_overlay_changed)
        overlay_layout.addWidget(self.chk_resistance)

        left_layout.addWidget(self.overlay_group)

        # 坐标轴缩放选项组
        self.scale_group = QGroupBox(self._tr("group_axis_scale"))
        scale_layout = QVBoxLayout(self.scale_group)

        # 左侧Y轴缩放
        self.left_label = QLabel(self._tr("label_left_axis"))
        self.left_label.setStyleSheet("font-weight: bold; color: #0969da;")
        scale_layout.addWidget(self.left_label)

        self.combo_left_scale = QComboBox()
        self.combo_left_scale.addItems([self._tr("scale_linear"), self._tr("scale_log")])
        self.combo_left_scale.currentIndexChanged.connect(self._on_scale_changed)
        scale_layout.addWidget(self.combo_left_scale)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #d0d7de;")
        scale_layout.addWidget(separator)

        # 右侧Y轴缩放
        self.right_label = QLabel(self._tr("label_right_axis"))
        self.right_label.setStyleSheet("font-weight: bold; color: #6e7781;")
        scale_layout.addWidget(self.right_label)

        self.combo_right_scale = QComboBox()
        self.combo_right_scale.addItems([self._tr("scale_linear"), self._tr("scale_log")])
        self.combo_right_scale.setEnabled(False)  # 默认不可用
        self.combo_right_scale.currentIndexChanged.connect(self._on_scale_changed)
        scale_layout.addWidget(self.combo_right_scale)

        left_layout.addWidget(self.scale_group)

        # 添加弹性空间
        left_layout.addStretch()

        # 添加左侧面板到主布局
        main_layout.addWidget(left_panel)

        # 右侧内容区域
        right_content = QWidget()
        right_layout = QVBoxLayout(right_content)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # 使用分割器
        splitter = QSplitter(Qt.Vertical)

        # 绘图区域
        self.plot_canvas = PlotCanvas()
        splitter.addWidget(self.plot_canvas)

        # 数据表格
        self.data_table = QTableWidget()
        self.data_table.setColumnCount(10)
        self.data_table.setHorizontalHeaderLabels([
            self._tr("col_round"), self._tr("col_peak_i"), self._tr("col_idle_i"), self._tr("col_response"),
            self._tr("col_rise_time"), self._tr("col_fall_time"),
            self._tr("col_rise_imax"), self._tr("col_imin"), self._tr("col_fall_imax"), self._tr("col_fall_imin")
        ])
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.data_table.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setEditTriggers(QTableWidget.NoEditTriggers)
        splitter.addWidget(self.data_table)

        # 设置分割比例
        splitter.setSizes([600, 200])
        right_layout.addWidget(splitter)

        # 底部按钮栏
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.btn_save_svg = QPushButton(self._tr("btn_save_svg"))
        self.btn_save_svg.clicked.connect(self._save_svg)
        button_layout.addWidget(self.btn_save_svg)

        self.btn_save_png = QPushButton(self._tr("btn_save_png"))
        self.btn_save_png.clicked.connect(self._save_png)
        button_layout.addWidget(self.btn_save_png)

        self.btn_copy_clipboard = QPushButton(self._tr("btn_copy"))
        self.btn_copy_clipboard.clicked.connect(self._copy_to_clipboard)
        button_layout.addWidget(self.btn_copy_clipboard)

        right_layout.addLayout(button_layout)

        # 添加右侧内容到主布局
        main_layout.addWidget(right_content)

    def _setup_menu(self):
        """设置菜单栏"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu(self._tr("menu_file"))

        open_action = QAction(self._tr("menu_open"), self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._open_file)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        exit_action = QAction(self._tr("menu_exit"), self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 设置菜单
        settings_menu = menubar.addMenu(self._tr("menu_settings"))

        settings_action = QAction(self._tr("menu_preferences"), self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self._show_settings)
        settings_menu.addAction(settings_action)

    def _setup_status_bar(self):
        """设置状态栏"""
        self.statusBar().showMessage(self._tr("status_ready"))

    def _apply_gui_scale(self, scale: float):
        """应用GUI缩放比例"""
        app: QApplication = QApplication.instance()  # type: ignore[assignment]
        font = QFont("Segoe UI", int(11 * scale))
        app.setFont(font)
        scaled_style = FLAT_STYLE.replace("font-size: 11px", f"font-size: {int(11 * scale)}px")
        app.setStyleSheet(scaled_style)

    def _on_overlay_changed(self):
        """数据叠加选项改变时的处理"""
        sender = self.sender()
        if sender == self.chk_voltage:
            if self.chk_voltage.isChecked():
                self.chk_resistance.setEnabled(False)
                self.combo_right_scale.setEnabled(True)
                self.right_label.setStyleSheet("font-weight: bold; color: #c92a2a;")
            else:
                self.chk_resistance.setEnabled(True)
                if not self.chk_resistance.isChecked():
                    self.combo_right_scale.setEnabled(False)
                    self.right_label.setStyleSheet("font-weight: bold; color: #6e7781;")
        elif sender == self.chk_resistance:
            if self.chk_resistance.isChecked():
                self.chk_voltage.setEnabled(False)
                self.combo_right_scale.setEnabled(True)
                self.right_label.setStyleSheet("font-weight: bold; color: #862e9c;")
            else:
                self.chk_voltage.setEnabled(True)
                if not self.chk_voltage.isChecked():
                    self.combo_right_scale.setEnabled(False)
                    self.right_label.setStyleSheet("font-weight: bold; color: #6e7781;")

        # 重新绘制
        self._refresh_plot()

    def _on_scale_changed(self):
        """坐标轴缩放改变时的处理"""
        self._refresh_plot()

    def _refresh_plot(self):
        """刷新绘图"""
        if self.analysis_result:
            settings = self.settings_manager.get_settings()
            show_voltage = self.chk_voltage.isChecked()
            show_resistance = self.chk_resistance.isChecked()
            left_scale = 'log' if self.combo_left_scale.currentIndex() == 1 else 'linear'
            right_scale = 'log' if self.combo_right_scale.currentIndex() == 1 else 'linear'

            self.plot_canvas.plot_data(
                self.analysis_result, settings,
                show_voltage=show_voltage,
                show_resistance=show_resistance,
                left_scale=left_scale,
                right_scale=right_scale
            )

    def _open_file(self):
        """打开文件对话框"""
        last_dir = self.settings_manager.get_settings().last_open_dir
        filepath, _ = QFileDialog.getOpenFileName(
            self, self._tr("dlg_open_title"), last_dir,
            self._tr("dlg_open_filter")
        )

        if filepath:
            # 更新上次打开的目录
            self.settings_manager.update_last_open_dir(str(Path(filepath).parent))
            self.current_file_path = filepath
            self._run_analysis(filepath)

    def _run_analysis(self, filepath: str):
        """运行分析"""
        self.statusBar().showMessage(self._tr("status_analyzing", filename=Path(filepath).name))
        QApplication.processEvents()

        try:
            settings = self.settings_manager.get_settings()
            self.analysis_result = run_analysis(
                filepath,
                rise_lower_percent=settings.analysis.rise_lower_percent,
                rise_upper_percent=settings.analysis.rise_upper_percent,
                fall_upper_percent=settings.analysis.fall_upper_percent,
                fall_lower_percent=settings.analysis.fall_lower_percent,
                prominence=settings.analysis.prominence,
                distance=settings.analysis.distance,
                baseline_fraction=settings.analysis.baseline_fraction,
                peak_threshold=settings.analysis.peak_threshold,
            )

            # 更新绘图
            self._refresh_plot()

            # 更新表格
            self._update_table()

            self.statusBar().showMessage(
                self._tr("status_analyzed", filename=Path(filepath).name, n=len(self.analysis_result.rounds))
            )
        except Exception as e:
            QMessageBox.critical(self, self._tr("msg_error"), self._tr("msg_analysis_failed", error=str(e)))
            self.statusBar().showMessage(self._tr("status_failed"))

    def _update_table(self):
        """更新数据表格"""
        if not self.analysis_result:
            return

        rounds = self.analysis_result.rounds
        self.data_table.setRowCount(len(rounds))

        # 自动选择电流单位
        max_current = np.max(self.analysis_result.currents)
        current_unit, current_scale = get_current_unit_and_scale(max_current)

        # 更新表头，添加单位信息
        self.data_table.setHorizontalHeaderLabels([
            self._tr("col_round"),
            f'{self._tr("col_peak_i")} ({current_unit})',
            f'{self._tr("col_idle_i")} ({current_unit})',
            self._tr("col_response"),
            self._tr("col_rise_time"),
            self._tr("col_fall_time"),
            f'{self._tr("col_rise_imax")} ({current_unit})',
            f'{self._tr("col_imin")} ({current_unit})',
            f'{self._tr("col_fall_imax")} ({current_unit})',
            f'{self._tr("col_fall_imin")} ({current_unit})'
        ])

        for i, r in enumerate(rounds):
            self.data_table.setItem(i, 0, QTableWidgetItem(str(r.round_num)))
            self.data_table.setItem(i, 1, QTableWidgetItem(f"{r.peak_current * current_scale:.4f}"))

            # Idle I = I_min（上升沿的基线截断均值）
            idle_str = f"{r.rise.min_val * current_scale:.4f}" if r.rise else "N/A"
            self.data_table.setItem(i, 2, QTableWidgetItem(idle_str))

            # Response = Peak I / Idle I
            response_str = f"{r.rise.ratio:.2f}" if r.rise else "N/A"
            self.data_table.setItem(i, 3, QTableWidgetItem(response_str))

            # Rise Time / Fall Time
            rise_time_str = f"{r.rise.transition_time:.4f}" if r.rise else "N/A"
            fall_time_str = f"{r.fall.transition_time:.4f}" if r.fall else "N/A"
            self.data_table.setItem(i, 4, QTableWidgetItem(rise_time_str))
            self.data_table.setItem(i, 5, QTableWidgetItem(fall_time_str))

            # Rise I_max / Rise I_min（用户定义百分比位置的电流值）
            if r.rise:
                self.data_table.setItem(i, 6, QTableWidgetItem(f"{r.rise.i_at_upper * current_scale:.4f}"))
                self.data_table.setItem(i, 7, QTableWidgetItem(f"{r.rise.i_at_lower * current_scale:.4f}"))
            else:
                self.data_table.setItem(i, 6, QTableWidgetItem("N/A"))
                self.data_table.setItem(i, 7, QTableWidgetItem("N/A"))

            # Fall I_max / Fall I_min
            if r.fall:
                self.data_table.setItem(i, 8, QTableWidgetItem(f"{r.fall.i_at_upper * current_scale:.4f}"))
                self.data_table.setItem(i, 9, QTableWidgetItem(f"{r.fall.i_at_lower * current_scale:.4f}"))
            else:
                self.data_table.setItem(i, 8, QTableWidgetItem("N/A"))
                self.data_table.setItem(i, 9, QTableWidgetItem("N/A"))

    def _show_settings(self):
        """显示设置对话框"""
        dialog = SettingsDialog(self.settings_manager.get_settings(), self.lang, self)
        if dialog.exec() == QDialog.Accepted:
            new_settings = dialog.get_settings()
            analysis_changed = dialog.get_analysis_changed()

            # 更新设置
            self.settings_manager.update_analysis(**new_settings['analysis'])
            self.settings_manager.update_plot(**new_settings['plot'])
            self.settings_manager.update_gui_scale(new_settings['gui_scale'])
            self.settings_manager.update_language(new_settings['language'])

            # 应用语言切换
            old_lang = self.lang.language
            new_lang = new_settings['language']
            if new_lang != old_lang:
                self.lang.set_language(new_lang)
                self._retranslate_ui()

            # 重新加载设置
            self.settings_manager.load()

            # 应用新的GUI缩放
            scale = new_settings['gui_scale']
            self._apply_gui_scale(scale)

            if self.analysis_result:
                if analysis_changed and self.current_file_path:
                    self._run_analysis(self.current_file_path)
                else:
                    self._refresh_plot()
                    self._update_table()

    def _retranslate_ui(self):
        """重新翻译所有UI文本"""
        self.setWindowTitle(self._tr("window_title"))
        self.statusBar().showMessage(self._tr("status_ready"))

        # 左侧面板
        self.overlay_group.setTitle(self._tr("group_overlay"))
        self.chk_voltage.setText(self._tr("chk_voltage"))
        self.chk_resistance.setText(self._tr("chk_resistance"))
        self.scale_group.setTitle(self._tr("group_axis_scale"))
        self.left_label.setText(self._tr("label_left_axis"))
        self.right_label.setText(self._tr("label_right_axis"))

        # 按钮
        self.btn_save_svg.setText(self._tr("btn_save_svg"))
        self.btn_save_png.setText(self._tr("btn_save_png"))
        self.btn_copy_clipboard.setText(self._tr("btn_copy"))

        # 表格
        self.data_table.setHorizontalHeaderLabels([
            self._tr("col_round"), self._tr("col_peak_i"), self._tr("col_idle_i"), self._tr("col_response"),
            self._tr("col_rise_time"), self._tr("col_fall_time"),
            self._tr("col_rise_imax"), self._tr("col_imin"), self._tr("col_fall_imax"), self._tr("col_fall_imin")
        ])

        # 菜单栏需要重建
        self.menuBar().clear()
        self._setup_menu()

    def _save_svg(self):
        """保存为 SVG"""
        if not self.analysis_result:
            QMessageBox.warning(self, self._tr("msg_warning"), self._tr("msg_no_data_save"))
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, self._tr("dlg_save_svg_title"), "", self._tr("dlg_save_svg_filter")
        )
        if filepath:
            self.plot_canvas.get_figure().savefig(filepath, format='svg', bbox_inches='tight')
            self.statusBar().showMessage(self._tr("status_saved", filepath=filepath))

    def _save_png(self):
        """保存为 PNG"""
        if not self.analysis_result:
            QMessageBox.warning(self, self._tr("msg_warning"), self._tr("msg_no_data_save"))
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, self._tr("dlg_save_png_title"), "", self._tr("dlg_save_png_filter")
        )
        if filepath:
            self.plot_canvas.get_figure().savefig(filepath, format='png', dpi=150, bbox_inches='tight')
            self.statusBar().showMessage(self._tr("status_saved", filepath=filepath))

    def _copy_to_clipboard(self):
        """复制图片到剪贴板（PNG格式）"""
        if not self.analysis_result:
            QMessageBox.warning(self, self._tr("msg_warning"), self._tr("msg_no_data_copy"))
            return

        from PySide6.QtGui import QImage
        from io import BytesIO

        # 将 Figure 保存为 PNG 到内存
        buf = BytesIO()
        self.plot_canvas.get_figure().savefig(buf, format='png', dpi=150, bbox_inches='tight')
        buf.seek(0)

        # 转换为 QImage 并复制到剪贴板
        image = QImage()
        image.loadFromData(buf.getvalue())

        clipboard = QApplication.clipboard()
        clipboard.setImage(image)

        self.statusBar().showMessage(self._tr("status_copied"))


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Gas Sensor Analyzer")
    app.setOrganizationName("GasSensor")

    # 设置应用样式
    app.setStyle("Fusion")

    # 加载保存的GUI缩放比例
    sm = SettingsManager()
    gui_scale = sm.get_settings().gui_scale
    scaled_style = FLAT_STYLE.replace("font-size: 11px", f"font-size: {int(11 * gui_scale)}px")
    app.setStyleSheet(scaled_style)

    # 设置全局字体
    font = QFont("Segoe UI", int(11 * gui_scale))
    app.setFont(font)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
