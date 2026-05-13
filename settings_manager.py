"""
设置管理模块。
负责加载、保存和管理应用程序设置。
"""

import json
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional


@dataclass
class AnalysisSettings:
    """分析参数设置"""
    rise_lower_percent: float = 10.0    # 上升沿下限百分比
    rise_upper_percent: float = 90.0    # 上升沿上限百分比
    fall_upper_percent: float = 90.0    # 下降沿上限百分比
    fall_lower_percent: float = 10.0    # 下降沿下限百分比
    prominence: float = 0.3             # 峰值检测显著性（log10空间）
    distance: int = 150                 # 相邻峰最小间距（采样点数）
    baseline_fraction: float = 0.01     # 基线阈值（占峰值的比例）
    peak_threshold: float = 0.0         # 峰值电流绝对阈值（A）


@dataclass
class PlotSettings:
    """绘图设置"""
    font_size_title: int = 16           # 标题字体大小
    font_size_axis_label: int = 15      # 轴标签字体大小
    font_size_tick: int = 12            # 刻度字体大小
    font_size_legend: int = 10          # 图例字体大小
    font_size_annotation: int = 10      # 标注字体大小
    title_bold: bool = True             # 标题是否加粗
    axis_label_bold: bool = True        # 轴标签是否加粗
    line_width_raw: float = 0.5         # 原始数据线宽
    line_width_smooth: float = 0.8      # 平滑数据线宽
    line_width_edge: float = 2.0        # 边沿线宽
    # 颜色设置
    color_raw: str = "#000000"          # 原始数据颜色
    color_smooth: str = "#1f77b4"       # 平滑数据颜色
    color_rise: str = "#2ca02c"         # 上升沿颜色
    color_fall: str = "#d62728"         # 下降沿颜色
    color_peak: str = "#ff0000"         # 峰值标记颜色
    color_rise_marker: str = "#2ca02c"  # 上升沿交叉点标记颜色
    color_fall_marker: str = "#ff7f0e"  # 下降沿交叉点标记颜色
    color_voltage: str = "#c92a2a"      # 电压曲线颜色
    color_resistance: str = "#862e9c"   # 电阻曲线颜色


@dataclass
class AppSettings:
    """应用程序完整设置"""
    analysis: AnalysisSettings = None
    plot: PlotSettings = None
    last_open_dir: str = ""             # 上次打开的目录
    gui_scale: float = 1.5              # GUI缩放比例
    language: str = "zh"                # 界面语言
    b_channel: str = ""                 # B通道选择："" / "resistance" / "voltage"

    def __post_init__(self):
        if self.analysis is None:
            self.analysis = AnalysisSettings()
        if self.plot is None:
            self.plot = PlotSettings()


class SettingsManager:
    """设置管理器"""

    def __init__(self, settings_path: Optional[str] = None):
        if settings_path is None:
            # 默认在脚本同目录下
            self.settings_path = Path(__file__).parent / "settings.json"
        else:
            self.settings_path = Path(settings_path)
        self.settings = AppSettings()
        self.load()

    def load(self) -> AppSettings:
        """从文件加载设置，如果文件不存在则创建默认设置"""
        if self.settings_path.exists():
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # 重建设置对象
                if 'analysis' in data:
                    self.settings.analysis = AnalysisSettings(**data['analysis'])
                if 'plot' in data:
                    self.settings.plot = PlotSettings(**data['plot'])
                if 'last_open_dir' in data:
                    self.settings.last_open_dir = data['last_open_dir']
                if 'gui_scale' in data:
                    self.settings.gui_scale = data['gui_scale']
                if 'language' in data:
                    self.settings.language = data['language']
                if 'b_channel' in data:
                    self.settings.b_channel = data['b_channel']
            except (json.JSONDecodeError, TypeError) as e:
                print(f"警告: 设置文件格式错误，使用默认设置: {e}")
                self.settings = AppSettings()
        else:
            # 文件不存在，创建默认设置
            self.save()
        return self.settings

    def save(self):
        """保存当前设置到文件"""
        data = {
            'analysis': asdict(self.settings.analysis),
            'plot': asdict(self.settings.plot),
            'last_open_dir': self.settings.last_open_dir,
            'gui_scale': self.settings.gui_scale,
            'language': self.settings.language,
            'b_channel': self.settings.b_channel,
        }
        with open(self.settings_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def get_settings(self) -> AppSettings:
        """获取当前设置"""
        return self.settings

    def update_analysis(self, **kwargs):
        """更新分析设置"""
        for key, value in kwargs.items():
            if hasattr(self.settings.analysis, key):
                setattr(self.settings.analysis, key, value)
        self.save()

    def update_plot(self, **kwargs):
        """更新绘图设置"""
        for key, value in kwargs.items():
            if hasattr(self.settings.plot, key):
                setattr(self.settings.plot, key, value)
        self.save()

    def update_gui_scale(self, scale: float):
        """更新GUI缩放比例"""
        self.settings.gui_scale = scale
        self.save()

    def update_detection(self, **kwargs):
        """更新峰值检测设置"""
        for key, value in kwargs.items():
            if hasattr(self.settings.analysis, key):
                setattr(self.settings.analysis, key, value)
        self.save()

    def update_last_open_dir(self, directory: str):
        """更新上次打开的目录"""
        self.settings.last_open_dir = directory
        self.save()

    def update_language(self, language: str):
        """更新界面语言"""
        self.settings.language = language
        self.save()

    def update_b_channel(self, b_channel: str):
        """更新B通道选择"""
        self.settings.b_channel = b_channel
        self.save()
