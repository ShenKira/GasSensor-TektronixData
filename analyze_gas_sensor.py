"""
气体传感器实验数据分析脚本（命令行版本）。
读取 testdata.xlsx，按轮次切分电流数据，分析上升沿/下降沿特征并绘图。
"""

import sys
import io
import warnings

import numpy as np
import matplotlib.pyplot as plt
import matplotlib

from settings_manager import SettingsManager
from analyze_core import run_analysis

# Windows 控制台 UTF-8 输出
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# matplotlib 字体配置
matplotlib.rcParams['font.family'] = 'serif'
matplotlib.rcParams['font.serif'] = ['Times New Roman']
matplotlib.rcParams['mathtext.fontset'] = 'stix'
matplotlib.rcParams['axes.unicode_minus'] = False
warnings.filterwarnings('ignore', message=r'.*Glyph.*missing from font.*')


def print_results(results, settings):
    """打印逐轮次分析结果。"""
    analysis = settings.analysis
    sep = "=" * 130
    dash = "-" * 130

    print(f"\n{sep}")
    print(f"{'PER-ROUND ANALYSIS RESULTS':^130}")
    print(sep)

    hdr = (f"{'Round':>5}  {'Peak Time(s)':>12}  {'Peak I(A)':>14}  "
           f"{'Rise I_max(A)':>14}  {'Rise I_min(A)':>14}  {'I_max/I_min':>12}  "
           f"{f'{analysis.rise_lower_percent:.0f}-{analysis.rise_upper_percent:.0f}% Rise(s)':>16}  "
           f"{f'{analysis.fall_upper_percent:.0f}-{analysis.fall_lower_percent:.0f}% Fall(s)':>16}")
    print(hdr)
    print(dash)

    for r in results:
        rise = r.rise
        fall = r.fall
        rise_ratio_s = f"{rise.ratio:.2f}" if rise else "N/A"
        rise_time_s = f"{rise.transition_time:.4f}" if rise else "N/A"
        fall_time_s = f"{fall.transition_time:.4f}" if fall else "N/A"
        rise_max_s = f"{rise.max_val:.6e}" if rise else "N/A"
        rise_min_s = f"{rise.min_val:.6e}" if rise else "N/A"

        print(f"{r.round_num:>5}  {r.peak_time:>12.2f}  {r.peak_current:>14.6e}  "
              f"{rise_max_s:>14}  {rise_min_s:>14}  {rise_ratio_s:>12}  "
              f"{rise_time_s:>16}  {fall_time_s:>16}")

    print(sep)

    # 详细逐轮分析
    print(f"\n{'DETAILED PER-ROUND BREAKDOWN':^130}")
    print(sep)

    for r in results:
        rise = r.rise
        fall = r.fall
        print(f"\n  Round #{r.round_num}")
        print(f"    Segment indices   : start={r.start_idx}, peak={r.peak_idx}, end={r.end_idx}")
        print(f"    Segment time span : {r.peak_time:.2f} s (peak)")

        if rise:
            print(f"    --- Rising Edge ---")
            print(f"      I_max           : {rise.max_val:.6e} A")
            print(f"      I_min           : {rise.min_val:.6e} A")
            print(f"      I_max / I_min   : {rise.ratio:.2f}")
            print(f"      t({analysis.rise_lower_percent:.0f}%)          : {rise.t_lower:.4f} s")
            print(f"      t({analysis.rise_upper_percent:.0f}%)          : {rise.t_upper:.4f} s")
            print(f"      {analysis.rise_lower_percent:.0f}% -> {analysis.rise_upper_percent:.0f}% time : {rise.transition_time:.4f} s")
        else:
            print(f"    --- Rising Edge : N/A ---")

        if fall:
            print(f"    --- Falling Edge ---")
            print(f"      I_max           : {fall.max_val:.6e} A")
            print(f"      I_min           : {fall.min_val:.6e} A")
            print(f"      t({analysis.fall_upper_percent:.0f}%)          : {fall.t_upper:.4f} s")
            print(f"      t({analysis.fall_lower_percent:.0f}%)          : {fall.t_lower:.4f} s")
            print(f"      {analysis.fall_upper_percent:.0f}% -> {analysis.fall_lower_percent:.0f}% time : {fall.transition_time:.4f} s")
        else:
            print(f"    --- Falling Edge : N/A ---")

    print(f"\n{sep}")


def plot_results(result, settings):
    """绘制分析结果。"""
    times = result.times
    currents = result.currents
    currents_smooth = result.currents_smooth
    rounds = result.rounds
    plot_settings = settings.plot
    analysis_settings = settings.analysis

    fig, ax = plt.subplots(figsize=(14, 6))

    font_family = 'Times New Roman'

    # 全局电流-时间曲线
    ax.plot(times, currents, 'k-', linewidth=plot_settings.line_width_raw, alpha=0.4, label='Raw data')
    ax.plot(times, currents_smooth, 'b-', linewidth=plot_settings.line_width_smooth, alpha=0.7, label='Smoothed')

    # 标记每一轮的上升沿和下降沿
    colors_rise = plt.cm.Greens(np.linspace(0.4, 0.9, len(rounds)))
    colors_fall = plt.cm.Reds(np.linspace(0.4, 0.9, len(rounds)))

    for i, r in enumerate(rounds):
        start, peak, end = r.start_idx, r.peak_idx, r.end_idx

        # 使用平滑数据，保证与峰值检测/分析结果的一致性
        ax.plot(
            times[start:peak + 1], currents_smooth[start:peak + 1],
            color=colors_rise[i], linewidth=plot_settings.line_width_edge,
        )
        ax.plot(
            times[peak:end + 1], currents_smooth[peak:end + 1],
            color=colors_fall[i], linewidth=plot_settings.line_width_edge,
        )

        ax.plot(times[peak], currents_smooth[peak], 'v', color='red', markersize=6)
        ax.annotate(
            f'#{r.round_num}',
            xy=(times[peak], currents_smooth[peak]),
            xytext=(5, 10), textcoords='offset points',
            fontsize=plot_settings.font_size_annotation,
            color='darkred', fontfamily=font_family,
        )

        if r.rise:
            for key, marker in [('t_lower', 'o'), ('t_upper', 's')]:
                t_val = getattr(r.rise, key)
                idx = np.argmin(np.abs(times - t_val))
                ax.plot(t_val, currents_smooth[idx], marker, color='green', markersize=5)
        if r.fall:
            for key, marker in [('t_upper', 'o'), ('t_lower', 's')]:
                t_val = getattr(r.fall, key)
                idx = np.argmin(np.abs(times - t_val))
                ax.plot(t_val, currents_smooth[idx], marker, color='orange', markersize=5)

    # 图例
    legend_elements = [
        Line2D([0], [0], color='green', linewidth=2, label='Rising edge'),
        Line2D([0], [0], color='red', linewidth=2, label='Falling edge'),
        Line2D([0], [0], marker='v', color='red', linestyle='None', markersize=6, label='Peak'),
        Line2D([0], [0], marker='o', color='green', linestyle='None', markersize=5,
               label=f'{analysis_settings.rise_lower_percent:.0f}% point (rise)'),
        Line2D([0], [0], marker='s', color='green', linestyle='None', markersize=5,
               label=f'{analysis_settings.rise_upper_percent:.0f}% point (rise)'),
    ]
    ax.legend(handles=legend_elements, loc='upper right',
              fontsize=plot_settings.font_size_legend, prop={'family': font_family})

    ax.set_xlabel('Time (s)', fontsize=plot_settings.font_size_axis_label,
                  fontweight='bold' if plot_settings.axis_label_bold else 'normal',
                  fontfamily=font_family)
    ax.set_ylabel('Current (A)', fontsize=plot_settings.font_size_axis_label,
                  fontweight='bold' if plot_settings.axis_label_bold else 'normal',
                  fontfamily=font_family)
    ax.set_title('Gas Sensor Experiment — Current vs Time (per-round annotation)',
                 fontsize=plot_settings.font_size_title,
                 fontweight='bold' if plot_settings.title_bold else 'normal',
                 fontfamily=font_family)
    ax.grid(True, which='both', alpha=0.3)

    ax.tick_params(axis='both', labelsize=plot_settings.font_size_tick)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontfamily('Times New Roman')

    plt.tight_layout()
    plt.show()


def main():
    from matplotlib.lines import Line2D

    filepath = 'testdata.xlsx'
    print(f"Loading {filepath} ...")

    # 加载设置
    settings_manager = SettingsManager()
    settings = settings_manager.get_settings()

    # 运行分析
    result = run_analysis(
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

    print(f"  {len(result.times)} data points, time range {result.times[0]:.2f} ~ {result.times[-1]:.2f} s")
    print(f"\nDetected {len(result.rounds)} experiment rounds")

    # 打印结果
    print_results(result.rounds, settings)

    # 绘图
    plot_results(result, settings)


if __name__ == '__main__':
    main()
