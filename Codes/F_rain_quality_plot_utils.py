"""Shared precipitation color utilities for rain/quality visualizations."""

from __future__ import annotations

import numpy as np
from matplotlib import colors as mcolors


RAIN_COLORSCALE = [
    [0.00, "#FFFFFF"],
    [0.10, "#D7F0FA"],
    [0.25, "#73BFE2"],
    [0.45, "#2B8CBE"],
    [0.62, "#41AB5D"],
    [0.78, "#FDD049"],
    [0.90, "#F46D43"],
    [1.00, "#7F0000"],
]


def compute_rain_cap(values, percentile: float = 99, minimum: float = 1.0) -> float:
    """Return a robust cap for precipitation colors."""
    array = np.asarray(values, dtype=float)
    positive = array[np.isfinite(array) & (array > 0)]
    if positive.size == 0:
        return float(minimum)
    return max(float(np.nanpercentile(positive, percentile)), float(minimum))


def sqrt_scale_rain(value: float, rain_cap: float) -> float:
    """Scale precipitation to 0-1 with the same square-root transform as Stage 9."""
    if not np.isfinite(value) or value <= 0:
        return 0.0
    return float(np.sqrt(min(float(value), float(rain_cap)) / float(rain_cap)))


def interpolate_colorscale(scaled_value: float, colorscale=None):
    """Return an RGBA color from a Plotly-like colorscale."""
    colorscale = RAIN_COLORSCALE if colorscale is None else colorscale
    scaled_value = min(max(float(scaled_value), 0.0), 1.0)

    for (left_pos, left_color), (right_pos, right_color) in zip(colorscale[:-1], colorscale[1:]):
        if scaled_value <= right_pos:
            fraction = 0 if right_pos == left_pos else (scaled_value - left_pos) / (right_pos - left_pos)
            return tuple(
                (1 - fraction) * left + fraction * right
                for left, right in zip(mcolors.to_rgba(left_color), mcolors.to_rgba(right_color))
            )
    return mcolors.to_rgba(colorscale[-1][1])


def rain_color_from_value(value: float, rain_cap: float):
    """Map a precipitation value to the shared rain-grid RGBA color."""
    if not np.isfinite(value):
        return mcolors.to_rgba("#f2f2f2")
    if value <= 0:
        return mcolors.to_rgba("#FFFFFF")
    return interpolate_colorscale(sqrt_scale_rain(value, rain_cap))
