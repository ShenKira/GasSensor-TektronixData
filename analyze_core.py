"""
核心分析模块。
包含数据加载、平滑、轮次切分和边沿分析逻辑。
"""

import numpy as np
import openpyxl
from scipy.signal import savgol_filter, find_peaks
from dataclasses import dataclass
from typing import List, Optional, Tuple


@dataclass
class EdgeInfo:
    """边沿信息"""
    max_val: float
    min_val: float
    ratio: float
    t_lower: float          # 下限百分比对应的时间
    t_upper: float          # 上限百分比对应的时间
    transition_time: float
    i_at_lower: float       # t_lower 时刻的电流值
    i_at_upper: float       # t_upper 时刻的电流值


@dataclass
class RoundResult:
    """单轮分析结果"""
    round_num: int
    peak_time: float
    peak_current: float
    rise: Optional[EdgeInfo]
    fall: Optional[EdgeInfo]
    start_idx: int
    peak_idx: int
    end_idx: int


@dataclass
class AnalysisResult:
    """完整分析结果"""
    times: np.ndarray
    voltages: np.ndarray
    currents: np.ndarray
    resistances: np.ndarray
    currents_smooth: np.ndarray
    rounds: List[RoundResult]
    file_path: str


def load_data(filepath: str) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """读取 xlsx 文件，返回 (时间, 电压, 电流, 电阻) 数组。"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    times, voltages, currents, resistances = [], [], [], []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        times.append(row[1])
        voltages.append(row[2])
        currents.append(row[3])
        resistances.append(row[4])

    return (
        np.array(times, dtype=float),
        np.array(voltages, dtype=float),
        np.array(currents, dtype=float),
        np.array(resistances, dtype=float),
    )


def smooth_current(currents: np.ndarray, window: int = 21, poly: int = 3) -> np.ndarray:
    """Savitzky-Golay 平滑，消除单点毛刺。"""
    window = min(window, len(currents))
    if window % 2 == 0:
        window -= 1
    smoothed = savgol_filter(currents, window_length=window, polyorder=poly)
    # 滤波器可能在极低值区域产生负值，截断到非负
    return np.clip(smoothed, 0, None)


def find_round_boundaries(
    times: np.ndarray,
    currents_smooth: np.ndarray,
    prominence: float = 0.3,
    distance: int = 150,
    baseline_fraction: float = 0.01,
    peak_threshold: float = 1e-7,
) -> List[Tuple[int, int, int]]:
    """
    识别每一轮实验的 (start_idx, peak_idx, end_idx)。
    """
    log_current = np.log10(np.clip(currents_smooth, 1e-20, None))

    peaks, _ = find_peaks(log_current, prominence=prominence, distance=distance)

    valid_peaks = [p for p in peaks if currents_smooth[p] > peak_threshold]

    if len(valid_peaks) == 0:
        return []

    rounds = []

    for i, pk in enumerate(valid_peaks):
        peak_val = currents_smooth[pk]
        threshold = peak_val * baseline_fraction

        start = 0
        for j in range(pk, -1, -1):
            if currents_smooth[j] <= threshold:
                start = j
                break

        right_limit = len(times) - 1
        if i + 1 < len(valid_peaks):
            right_limit = valid_peaks[i + 1]

        end = right_limit
        for j in range(pk, min(right_limit + 1, len(times))):
            if currents_smooth[j] <= threshold:
                end = j
                break

        rounds.append((start, pk, end))

    return rounds


def interpolate_crossing_time(
    times: np.ndarray,
    values: np.ndarray,
    idx_before: int,
    idx_after: int,
    threshold: float,
) -> float:
    """线性插值找到精确的交叉时间。"""
    if idx_after == idx_before:
        return times[idx_before]
    t0, t1 = times[idx_before], times[idx_after]
    v0, v1 = values[idx_before], values[idx_after]
    if v1 == v0:
        return (t0 + t1) / 2.0
    frac = (threshold - v0) / (v1 - v0)
    frac = np.clip(frac, 0.0, 1.0)
    return t0 + frac * (t1 - t0)


def _find_crossing(
    seg_t: np.ndarray,
    seg_i: np.ndarray,
    threshold: float,
    direction: str,
) -> Optional[float]:
    """在线段中找到第一次穿过 threshold 的精确时刻。"""
    for k in range(len(seg_i) - 1):
        if direction == 'up':
            if seg_i[k] <= threshold <= seg_i[k + 1]:
                return interpolate_crossing_time(seg_t, seg_i, k, k + 1, threshold)
        else:
            if seg_i[k] >= threshold >= seg_i[k + 1]:
                return interpolate_crossing_time(seg_t, seg_i, k, k + 1, threshold)
    return None


def _trimmed_mean(data: np.ndarray, trim_fraction: float = 0.25) -> float:
    """去除最高和最低各 trim_fraction 比例后取平均值。"""
    n = len(data)
    if n == 0:
        return 0.0
    k = int(n * trim_fraction)
    if k * 2 >= n:
        # 数据太少，直接取均值
        return float(np.mean(data))
    sorted_data = np.sort(data)
    trimmed = sorted_data[k:n - k]
    return float(np.mean(trimmed))


def analyze_edge(
    times: np.ndarray,
    currents_raw: np.ndarray,
    currents_smooth: np.ndarray,
    start_idx: int,
    end_idx: int,
    is_rising: bool,
    lower_percent: float = 10.0,
    upper_percent: float = 90.0,
) -> Optional[EdgeInfo]:
    """
    分析一段上升沿或下降沿。
    使用平滑数据确定交叉点和 min/max，保证一致性。
    上升沿的 I_min 取 start_idx 之前 24 点的截断均值。
    """
    seg_t = times[start_idx:end_idx + 1]
    seg_smooth = currents_smooth[start_idx:end_idx + 1]

    if len(seg_smooth) < 2:
        return None

    max_val = float(np.max(seg_smooth))

    # 上升沿：I_min 取上升开始前 24 点的截断均值（去头尾各25%）
    if is_rising:
        baseline_start = max(0, start_idx - 24)
        baseline_window = currents_smooth[baseline_start:start_idx]
        if len(baseline_window) >= 4:
            min_val = _trimmed_mean(baseline_window, trim_fraction=0.25)
        elif len(baseline_window) > 0:
            min_val = float(np.mean(baseline_window))
        else:
            min_val = float(np.min(seg_smooth))
    else:
        min_val = float(np.min(seg_smooth))

    if max_val <= min_val:
        return None

    ratio = max_val / min_val if min_val > 0 else float('inf')

    level_lower = min_val + (lower_percent / 100.0) * (max_val - min_val)
    level_upper = min_val + (upper_percent / 100.0) * (max_val - min_val)

    if is_rising:
        t_lower = _find_crossing(seg_t, seg_smooth, level_lower, direction='up')
        t_upper = _find_crossing(seg_t, seg_smooth, level_upper, direction='up')
    else:
        t_upper = _find_crossing(seg_t, seg_smooth, level_upper, direction='down')
        t_lower = _find_crossing(seg_t, seg_smooth, level_lower, direction='down')

    if t_lower is None or t_upper is None:
        return None

    transition_time = abs(t_upper - t_lower)

    # 用平滑数据中最近邻索引获取交叉点处的电流值
    idx_lower = int(np.argmin(np.abs(times - t_lower)))
    idx_upper = int(np.argmin(np.abs(times - t_upper)))

    return EdgeInfo(
        max_val=max_val,
        min_val=min_val,
        ratio=ratio,
        t_lower=t_lower,
        t_upper=t_upper,
        transition_time=transition_time,
        i_at_lower=float(currents_smooth[idx_lower]),
        i_at_upper=float(currents_smooth[idx_upper]),
    )


def run_analysis(
    filepath: str,
    rise_lower_percent: float = 10.0,
    rise_upper_percent: float = 90.0,
    fall_upper_percent: float = 90.0,
    fall_lower_percent: float = 10.0,
    prominence: float = 0.3,
    distance: int = 150,
    baseline_fraction: float = 0.01,
    peak_threshold: float = 1e-7,
) -> AnalysisResult:
    """运行完整的分析流程。"""
    times, voltages, currents, resistances = load_data(filepath)
    currents_smooth = smooth_current(currents, window=21, poly=3)

    round_boundaries = find_round_boundaries(
        times, currents_smooth,
        prominence=prominence,
        distance=distance,
        baseline_fraction=baseline_fraction,
        peak_threshold=peak_threshold,
    )

    rounds = []
    for i, (start, peak, end) in enumerate(round_boundaries):
        # 上升沿分析：使用平滑数据
        rise_info = analyze_edge(
            times, currents, currents_smooth, start, peak, is_rising=True,
            lower_percent=rise_lower_percent,
            upper_percent=rise_upper_percent,
        )
        # 下降沿分析：使用平滑数据
        fall_info = analyze_edge(
            times, currents, currents_smooth, peak, end, is_rising=False,
            lower_percent=fall_lower_percent,
            upper_percent=fall_upper_percent,
        )
        rounds.append(RoundResult(
            round_num=i + 1,
            peak_time=times[peak],
            peak_current=currents_smooth[peak],
            rise=rise_info,
            fall=fall_info,
            start_idx=start,
            peak_idx=peak,
            end_idx=end,
        ))

    return AnalysisResult(
        times=times,
        voltages=voltages,
        currents=currents,
        resistances=resistances,
        currents_smooth=currents_smooth,
        rounds=rounds,
        file_path=filepath,
    )
