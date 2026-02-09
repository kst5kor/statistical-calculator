import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import scipy.stats as stats
import math
import plotly.graph_objects as go
import io
import datetime
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Fill, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# --- Statistical Calculator Logic ---
# Ported from the 'statisticalCalculator' JavaScript object
class StatisticalCalculator:
    def erf(self, x):
        return math.erf(x)

    def standard_normal_cdf(self, z):
        return 0.5 * (1 + self.erf(z / math.sqrt(2)))

    def get_critical_value(self, cl, type):
        alpha = 1 - (cl / 100.0)
        if type == "Two-Sided":
            return stats.norm.ppf(1 - alpha / 2)
        elif type == "Upper-Sided":
            return stats.norm.ppf(1 - alpha)
        else:  # Lower-Sided
            return stats.norm.ppf(
                alpha
            )  # Note: This will be negative, handle in CI calculation

    def validate(self, params):
        LSL = params.get("lsl")
        USL = params.get("usl")
        s = params.get("s")
        n_samples = params.get("n_samples")
        confidenceLevel = params.get("confidence_level")
        distribution = params.get("distribution")
        Tm = params.get("tm")
        importedData = params.get("importedData", [])
        mode = params.get("mode")

        try:
            if any(v is None for v in [LSL, USL, Tm]):
                return "All specification values (Tm, LSL, USL) must be valid numbers."
            if USL <= LSL:
                return "USL must be greater than LSL."
            if mode == "manual" and (s is None or s < 0):
                return "Standard Deviation must be zero or positive for manual input."
            if n_samples is None or n_samples < 2:
                return "Sample Size (n) must be at least 2."
            if (
                confidenceLevel is None
                or confidenceLevel <= 0
                or confidenceLevel >= 100
            ):
                return "Confidence Level must be between 0 and 100."
            if distribution == "Lognormal" and (LSL <= 0 or USL <= 0 or Tm <= 0):
                return "For Lognormal distribution, all specification limits and the target mean must be positive."
            if mode == "import" and importedData:
                if len(importedData) < 2:
                    return "Imported data must contain at least 2 valid numeric points."
                if distribution == "Lognormal" and any(d <= 0 for d in importedData):
                    return "Lognormal distribution requires all imported data points to be positive."
            elif mode == "import" and not importedData:
                return "Import mode selected, but no data found in Data Worksheet."
            if mode == "import" and s is None:
                return "Could not calculate Standard Deviation from imported data. Check data format."
            if mode == "import" and s < 0:
                return "Calculated Standard Deviation from imported data cannot be negative."
            return None
        except Exception as e:
            return f"Validation error: {e}"

    class Normal:
        def pdf(self, x, mean, stdDev):
            if stdDev <= 0 or not np.isfinite(stdDev):
                return 0
            return (1 / (stdDev * np.sqrt(2 * np.pi))) * np.exp(
                -0.5 * ((x - mean) / stdDev) ** 2
            )

        def calculate(self, params):
            x_bar, s, USL, LSL, Tm = (
                params["x_bar"],
                params["s"],
                params["usl"],
                params["lsl"],
                params["tm"],
            )
            if s == 0:
                prob_above = 1 if x_bar > USL else 0
                prob_below = 1 if x_bar < LSL else 0
                prob_below_target = 1 if x_bar < Tm else 0
                cpk_s0 = np.inf if LSL <= x_bar <= USL else -np.inf
                return {
                    **params,
                    "T_drawing": USL - LSL,
                    "sixSigmaSpread": 0,
                    "Cp": np.inf,
                    "CpkCurrent": cpk_s0,
                    "prob_above": prob_above,
                    "prob_below": prob_below,
                    "prob_below_target": prob_below_target,
                }
            if s < 0 or not np.isfinite(s):
                return {
                    **params,
                    "T_drawing": np.nan,
                    "sixSigmaSpread": np.nan,
                    "Cp": np.nan,
                    "CpkCurrent": np.nan,
                    "prob_above": np.nan,
                    "prob_below": np.nan,
                    "prob_below_target": np.nan,
                }

            T_drawing = USL - LSL
            sixSigmaSpread = 6 * s
            Cp = T_drawing / sixSigmaSpread
            CpkCurrent = min((USL - x_bar) / (3 * s), (x_bar - LSL) / (3 * s))
            z_usl = (USL - x_bar) / s
            z_lsl = (LSL - x_bar) / s
            z_target = (Tm - x_bar) / s

            prob_above = 1 - stats.norm.cdf(z_usl)
            prob_below = stats.norm.cdf(z_lsl)
            prob_below_target = stats.norm.cdf(z_target)
            return {
                **params,
                "T_drawing": T_drawing,
                "sixSigmaSpread": sixSigmaSpread,
                "Cp": Cp,
                "CpkCurrent": CpkCurrent,
                "prob_above": prob_above,
                "prob_below": prob_below,
                "prob_below_target": prob_below_target,
            }

    class Lognormal:
        def pdf(self, x, mu_log, sigma_log):
            if x <= 0 or sigma_log <= 0 or not np.isfinite(sigma_log):
                return 0
            term1 = 1 / (x * sigma_log * np.sqrt(2 * np.pi))
            term2 = np.exp(-((np.log(x) - mu_log) ** 2) / (2 * sigma_log**2))
            return term1 * term2

        def calculate(self, params):
            x_bar, s, USL, LSL, Tm = (
                params["x_bar"],
                params["s"],
                params["usl"],
                params["lsl"],
                params["tm"],
            )

            if s == 0 or x_bar <= 0 or not np.isfinite(x_bar):
                prob_above = 1 if x_bar > USL else 0
                prob_below = 1 if x_bar < LSL else 0
                prob_below_target = 1 if x_bar < Tm else 0
                cpk_s0 = np.inf if LSL <= x_bar <= USL and x_bar > 0 else -np.inf
                return {
                    **params,
                    "T_drawing": np.nan,
                    "sixSigmaSpread": 0,
                    "Cp": np.inf,
                    "CpkCurrent": cpk_s0,
                    "prob_above": prob_above,
                    "prob_below": prob_below,
                    "prob_below_target": prob_below_target,
                    "mu_log": np.nan,
                    "sigma_log": 0,
                }
            if s < 0 or not np.isfinite(s):
                return {
                    **params,
                    "T_drawing": np.nan,
                    "sixSigmaSpread": np.nan,
                    "Cp": np.nan,
                    "CpkCurrent": np.nan,
                    "prob_above": np.nan,
                    "prob_below": np.nan,
                    "prob_below_target": np.nan,
                    "mu_log": np.nan,
                    "sigma_log": np.nan,
                }

            if LSL <= 0 or USL <= 0 or Tm <= 0:
                return {
                    **params,
                    "error": "LSL, USL, and Tm must be positive for Lognormal distribution.",
                }

            try:
                sigma_log_sq = np.log(1 + (s**2 / x_bar**2))
                sigma_log = np.sqrt(sigma_log_sq)
                mu_log = np.log(x_bar) - 0.5 * sigma_log_sq
            except ValueError:
                return {
                    **params,
                    "error": "Failed to calculate lognormal parameters. Check data.",
                }

            LSL_log, USL_log, Tm_log = np.log(LSL), np.log(USL), np.log(Tm)

            if sigma_log <= 0 or not np.isfinite(sigma_log):
                prob_above = 1 if x_bar > USL else 0
                prob_below = 1 if x_bar < LSL else 0
                prob_below_target = 1 if x_bar < Tm else 0
                cpk_s0 = np.inf if LSL <= x_bar <= USL else -np.inf
                return {
                    **params,
                    "T_drawing": np.nan,
                    "sixSigmaSpread": 0,
                    "Cp": np.inf,
                    "CpkCurrent": cpk_s0,
                    "prob_above": prob_above,
                    "prob_below": prob_below,
                    "prob_below_target": prob_below_target,
                    "mu_log": mu_log,
                    "sigma_log": sigma_log,
                }

            Cp = (USL_log - LSL_log) / (6 * sigma_log)
            CpkCurrent = min(
                (USL_log - mu_log) / (3 * sigma_log),
                (mu_log - LSL_log) / (3 * sigma_log),
            )

            z_usl_log = (USL_log - mu_log) / sigma_log
            z_lsl_log = (LSL_log - mu_log) / sigma_log
            z_target_log = (Tm_log - mu_log) / sigma_log

            prob_above = 1 - stats.norm.cdf(z_usl_log)
            prob_below = stats.norm.cdf(z_lsl_log)
            prob_below_target = stats.norm.cdf(z_target_log)

            return {
                **params,
                "T_drawing": USL - LSL,
                "sixSigmaSpread": 6 * s,
                "Cp": Cp,
                "CpkCurrent": CpkCurrent,
                "prob_above": prob_above,
                "prob_below": prob_below,
                "prob_below_target": prob_below_target,
                "mu_log": mu_log,
                "sigma_log": sigma_log,
            }

    def parse_raw_data(self, data_string):
        if not data_string:
            return []
        values = re.split(r"[\s,;\n]+", data_string.strip())
        return [float(v) for v in values if v and self.is_numeric(v)]

    def is_numeric(self, s):
        try:
            float(s)
            return True
        except (ValueError, TypeError):
            return False

    def calculate(self, inputs):
        params = {
            "tm": inputs.get("tm"),
            "lsl": inputs.get("lsl"),
            "usl": inputs.get("usl"),
            "target_index_value": inputs.get("target_index_value", 1.67),
            "target_index_type": inputs.get("target_index_type", "Cpk"),
            "confidence_level": inputs.get("confidence_level", 95.0),
            "distribution": inputs.get("distribution", "Normal"),
            "dp": inputs.get("decimal_places", 3),
            "hypothesis_type": inputs.get("hypothesis_type", "Two-Sided"),
            "mode": inputs.get("mode", "manual"),
            "measurement_name": inputs.get("measurement_name", "Unnamed") or "Unnamed",
        }

        if params["mode"] == "import":
            data = self.parse_raw_data(inputs.get("raw_data", ""))
            params["importedData"] = data
            if len(data) >= 2:
                params["n_samples"] = len(data)
                params["x_bar"] = np.mean(data)
                params["s"] = np.std(data, ddof=1) if len(data) > 1 else 0
            else:
                params["n_samples"] = len(data)
                params["x_bar"] = np.nan
                params["s"] = np.nan
                params["importedData"] = []
        else:
            params["x_bar"] = inputs.get("x_bar")
            params["s"] = inputs.get("s")
            params["n_samples"] = inputs.get("n_samples")
            params["importedData"] = []

        validationError = self.validate(params)
        if validationError:
            return {**params, "error": validationError}

        # Check for non-numeric essential values
        essential_keys = ["tm", "lsl", "usl", "x_bar", "s", "n_samples"]
        if any(not self.is_numeric(params[k]) for k in essential_keys):
            if params.get("s") == 0 and all(
                self.is_numeric(params[k])
                for k in ["tm", "lsl", "usl", "x_bar", "n_samples"]
            ):
                pass  # s=0 is a valid case
            else:
                return {
                    **params,
                    "error": "Essential inputs (Tm, LSL, USL, x_bar, s, n) must be valid numbers.",
                }

        results = {}
        if params["distribution"] == "Lognormal":
            results = self.Lognormal().calculate(params)
        else:
            results = self.Normal().calculate(params)

        if results.get("error"):
            return results

        results["shiftValue"] = results["tm"] - results["x_bar"]
        results["newToleranceTotal"] = (
            results["target_index_value"] * 6 * results["s"]
            if results["s"] > 0 and np.isfinite(results["target_index_value"])
            else (0 if results["s"] == 0 else np.nan)
        )
        results["eightSigmaSpread"] = 8 * results["s"] if results["s"] >= 0 else np.nan
        results["minus3s"] = results["x_bar"] - 3 * results["s"]
        results["plus3s"] = results["x_bar"] + 3 * results["s"]
        results["minus4s"] = results["x_bar"] - 4 * results["s"]
        results["plus4s"] = results["x_bar"] + 4 * results["s"]
        results["ppm_above"] = results.get("prob_above", np.nan) * 1e6
        results["ppm_below"] = results.get("prob_below", np.nan) * 1e6

        alpha = 1 - (results["confidence_level"] / 100)

        if results["n_samples"] >= 2 and results["s"] >= 0:
            std_error = (
                results["s"] / np.sqrt(results["n_samples"]) if results["s"] > 0 else 0
            )
            z_stat = (
                (results["x_bar"] - results["tm"]) / std_error
                if std_error > 0
                else (
                    0
                    if results["x_bar"] == results["tm"]
                    else np.inf * np.sign(results["x_bar"] - results["tm"])
                )
            )

            p_value = np.nan
            if not np.isfinite(z_stat):
                p_value = 0.0
            elif results["hypothesis_type"] == "Two-Sided":
                p_value = 2 * (1 - stats.norm.cdf(abs(z_stat)))
            elif results["hypothesis_type"] == "Upper-Sided":  # mu > Tm
                p_value = 1 - stats.norm.cdf(z_stat)
            else:  # Lower-Sided, mu < Tm
                p_value = stats.norm.cdf(z_stat)

            # Use more precise ppf function
            criticalValue_ppf = (
                abs(stats.norm.ppf(alpha / 2))
                if results["hypothesis_type"] == "Two-Sided"
                else abs(stats.norm.ppf(alpha))
            )
            marginOfError = criticalValue_ppf * std_error

            if results["hypothesis_type"] == "Two-Sided":
                results["ci_lower"] = results["x_bar"] - marginOfError
                results["ci_upper"] = results["x_bar"] + marginOfError
            elif (
                results["hypothesis_type"] == "Upper-Sided"
            ):  # Test is mu > Tm, CI is for mu
                results["ci_lower"] = results["x_bar"] - marginOfError  # One-sided CI
                results["ci_upper"] = np.inf
            else:  # Lower-Sided
                results["ci_lower"] = -np.inf
                results["ci_upper"] = results["x_bar"] + marginOfError  # One-sided CI

            results["hypothesisResult"] = {
                "z_stat": z_stat,
                "p_value": p_value,
                "alpha": alpha,
            }
        else:
            results["ci_lower"] = np.nan
            results["ci_upper"] = np.nan
            results["hypothesisResult"] = {
                "z_stat": np.nan,
                "p_value": np.nan,
                "alpha": alpha,
            }

        return {**results, "error": None}


# --- Plotting Logic ---
# Ported from 'plotManager'
class PlotManager:
    def generate_pdf_data(self, dist_type, params, x_min, x_max, points=200):
        x = np.linspace(x_min, x_max, points)
        y = np.zeros_like(x)

        calc = StatisticalCalculator()

        if dist_type == "Normal" and params.get("stdDev", 0) > 0:
            y = [calc.Normal().pdf(val, params["mean"], params["stdDev"]) for val in x]
        elif dist_type == "Lognormal" and params.get("sigma_log", 0) > 0:
            y = [
                calc.Lognormal().pdf(val, params["mu_log"], params["sigma_log"])
                for val in x
            ]

        return x, np.nan_to_num(y)

    def update_plots(self, results):
        LSL, USL, x_bar, s, Tm, target_index_value, dp = (
            results.get("lsl"),
            results.get("usl"),
            results.get("x_bar"),
            results.get("s"),
            results.get("tm"),
            results.get("target_index_value"),
            results.get("dp"),
        )
        ci_lower, ci_upper, confidenceLevel, distribution = (
            results.get("ci_lower"),
            results.get("ci_upper"),
            results.get("confidence_level"),
            results.get("distribution"),
        )
        mu_log, sigma_log, importedData = (
            results.get("mu_log"),
            results.get("sigma_log"),
            results.get("importedData", []),
        )

        cannot_plot = (
            any(not np.isfinite(v) for v in [LSL, USL, x_bar, Tm])
            or s < 0
            or not np.isfinite(s)
        )
        if cannot_plot:
            return None, None, None  # Return empty figures

        newToleranceTotal = results.get("newToleranceTotal", np.nan)
        newLSL = (
            Tm - (newToleranceTotal / 2) if np.isfinite(newToleranceTotal) else np.nan
        )
        newUSL = (
            Tm + (newToleranceTotal / 2) if np.isfinite(newToleranceTotal) else np.nan
        )

        data_min = (
            min(importedData)
            if importedData
            else (x_bar - 4.5 * s if s > 0 else x_bar - 1)
        )
        data_max = (
            max(importedData)
            if importedData
            else (x_bar + 4.5 * s if s > 0 else x_bar + 1)
        )

        x_points = [
            LSL,
            USL,
            newLSL,
            newUSL,
            x_bar,
            Tm,
            ci_lower,
            ci_upper,
            x_bar - 4.5 * s if s > 0 else None,
            x_bar + 4.5 * s if s > 0 else None,
            data_min,
            data_max,
        ]

        finite_x = [p for p in x_points if p is not None and np.isfinite(p)]

        if not finite_x:
            x_min, x_max = Tm - 1, Tm + 1
        else:
            raw_min, raw_max = min(finite_x), max(finite_x)
            range_val = raw_max - raw_min
            min_range = max(
                s * 0.5 if s > 0 else 0.1,
                abs(Tm - x_bar) or 0.1,
                (USL - LSL) * 0.1 or 0.1,
                0.1,
            )
            if range_val < min_range or range_val == 0:
                range_val = min_range
            x_min = raw_min - range_val * 0.2
            x_max = raw_max + range_val * 0.2

        pdf_data_before_x, pdf_data_before_y = [], []
        pdf_data_after_x, pdf_data_after_y = [], []
        max_pdf_y = 1

        if s > 0 and np.isfinite(s):
            if distribution == "Lognormal":
                pdf_params = {"mu_log": mu_log, "sigma_log": sigma_log}
                pdf_params_centered = {
                    "mu_log": np.log(Tm) - 0.5 * (sigma_log**2),
                    "sigma_log": sigma_log,
                }
            else:
                pdf_params = {"mean": x_bar, "stdDev": s}
                pdf_params_centered = {"mean": Tm, "stdDev": s}

            pdf_data_before_x, pdf_data_before_y = self.generate_pdf_data(
                distribution, pdf_params, x_min, x_max
            )
            pdf_data_after_x, pdf_data_after_y = self.generate_pdf_data(
                distribution, pdf_params_centered, x_min, x_max
            )

            valid_y = [
                y
                for y in np.concatenate((pdf_data_before_y, pdf_data_after_y))
                if np.isfinite(y) and y > 0
            ]
            if valid_y:
                max_pdf_y = max(valid_y) * 1.1

        layout_defaults = {
            "xaxis": {
                "title": "Value",
                "range": [x_min, x_max],
                "tickfont": {"size": 9},
            },
            "yaxis": {
                "title": "Density" if s > 0 else "",
                "tickformat": ".2f" if s > 0 else "",
                "fixedrange": True,
                "range": [0, max_pdf_y],
                "tickfont": {"size": 9},
                "showticklabels": s > 0,
            },
            "legend": {
                "font": {"size": 9},
                "x": 0.98,
                "xanchor": "right",
                "y": 0.98,
                "yanchor": "top",
                "bgcolor": "rgba(255,255,255,0.7)",
            },
            "margin": {"t": 40, "b": 40, "l": 50, "r": 20},
            "hovermode": "x unified",
        }

        # Plot 1: Current Process
        fig_before = go.Figure()
        if s > 0:
            fig_before.add_trace(
                go.Scatter(
                    x=pdf_data_before_x,
                    y=pdf_data_before_y,
                    mode="lines",
                    name=f"Current PDF (x̄={x_bar:.{dp}f})",
                    fill="tozeroy",
                    fillcolor="rgba(185, 28, 28, 0.1)",
                    line={"color": "#B91C1C", "width": 2},
                )
            )
        fig_before.add_trace(
            go.Scatter(
                x=[LSL, LSL],
                y=[0, max_pdf_y * 0.95],
                mode="lines",
                name="LSL",
                line={"color": "#047857", "dash": "dash", "width": 1.5},
            )
        )
        fig_before.add_trace(
            go.Scatter(
                x=[USL, USL],
                y=[0, max_pdf_y * 0.95],
                mode="lines",
                name="USL",
                line={"color": "#047857", "dash": "dash", "width": 1.5},
            )
        )
        fig_before.add_trace(
            go.Scatter(
                x=[x_bar, x_bar],
                y=[0, max_pdf_y * 0.9],
                mode="lines",
                name="Mean (x̄)",
                line={
                    "color": "#DC2626",
                    "width": 3 if s == 0 else 1.5,
                    "dash": "solid",
                },
            )
        )
        fig_before.add_trace(
            go.Scatter(
                x=[Tm, Tm],
                y=[0, max_pdf_y * 0.8],
                mode="lines",
                name="Target (Tm)",
                line={"color": "#4B5563", "dash": "dot", "width": 1.5},
            )
        )

        shapes_before = [
            {
                "type": "rect",
                "xref": "x",
                "yref": "paper",
                "x0": x_min,
                "y0": 0,
                "x1": LSL,
                "y1": 1,
                "fillcolor": "rgba(239, 68, 68, 0.15)",
                "line": {"width": 0},
                "layer": "below",
            },
            {
                "type": "rect",
                "xref": "x",
                "yref": "paper",
                "x0": USL,
                "y0": 0,
                "x1": x_max,
                "y1": 1,
                "fillcolor": "rgba(239, 68, 68, 0.15)",
                "line": {"width": 0},
                "layer": "below",
            },
        ]
        annotations_before = []
        if s > 0 and np.isfinite(ci_lower) and np.isfinite(ci_upper):
            shapes_before.append(
                {
                    "type": "line",
                    "xref": "x",
                    "yref": "paper",
                    "x0": ci_lower,
                    "y0": 0.05,
                    "x1": ci_upper,
                    "y1": 0.05,
                    "line": {"color": "#F97316", "width": 4},
                    "layer": "above",
                }
            )
            annotations_before.append(
                {
                    "x": (ci_lower + ci_upper) / 2,
                    "y": 0.02,
                    "xref": "x",
                    "yref": "paper",
                    "text": f"{confidenceLevel}% CI",
                    "showarrow": False,
                    "font": {"size": 9, "color": "#F97316"},
                    "yanchor": "top",
                }
            )

        ci_text = (
            f"[{ci_lower:.{dp}f}, {ci_upper:.{dp}f}]"
            if np.isfinite(ci_lower) and np.isfinite(ci_upper)
            else (
                f"[{ci_lower:.{dp}f}, +∞)"
                if np.isfinite(ci_lower)
                else f"(-∞, {ci_upper:.{dp}f}]"
                if np.isfinite(ci_upper)
                else ""
            )
        )
        title_before = f"1. Current Process Distribution {'& CI (' + ci_text + ')' if s > 0 and ci_text else ('(σ=0)' if s == 0 else '')}"
        fig_before.update_layout(
            **layout_defaults,
            title={"text": title_before, "font": {"size": 12}},
            shapes=shapes_before,
            annotations=annotations_before,
        )

        # Plot 2: Centered Process
        fig_after = go.Figure()
        if s > 0:
            fig_after.add_trace(
                go.Scatter(
                    x=pdf_data_after_x,
                    y=pdf_data_after_y,
                    mode="lines",
                    name=f"Centered PDF (at Tm={Tm:.{dp}f})",
                    fill="tozeroy",
                    fillcolor="rgba(0, 123, 197, 0.1)",
                    line={"color": "#007BC5", "width": 2},
                )
            )
        if np.isfinite(newLSL):
            fig_after.add_trace(
                go.Scatter(
                    x=[newLSL, newLSL],
                    y=[0, max_pdf_y * 0.95],
                    mode="lines",
                    name="Req. LSL",
                    line={"color": "#004A86", "width": 1.5, "dash": "dot"},
                )
            )
        if np.isfinite(newUSL):
            fig_after.add_trace(
                go.Scatter(
                    x=[newUSL, newUSL],
                    y=[0, max_pdf_y * 0.95],
                    mode="lines",
                    name="Req. USL",
                    line={"color": "#004A86", "width": 1.5, "dash": "dot"},
                )
            )

        fig_after.add_trace(
            go.Scatter(
                x=[Tm, Tm],
                y=[0, max_pdf_y * 0.9],
                mode="lines",
                name="Target (Tm)",
                line={
                    "color": "#007BC5",
                    "width": 3 if s == 0 else 1.5,
                    "dash": "solid",
                },
            )
        )
        fig_after.add_trace(
            go.Scatter(
                x=[LSL, LSL],
                y=[0, max_pdf_y * 0.95],
                mode="lines",
                name="Orig. LSL",
                line={"color": "rgba(4, 120, 87, 0.5)", "dash": "dash", "width": 1},
            )
        )
        fig_after.add_trace(
            go.Scatter(
                x=[USL, USL],
                y=[0, max_pdf_y * 0.95],
                mode="lines",
                name="Orig. USL",
                line={"color": "rgba(4, 120, 87, 0.5)", "dash": "dash", "width": 1},
            )
        )

        shapes_after = []
        if np.isfinite(newLSL) and np.isfinite(newUSL):
            shapes_after.append(
                {
                    "type": "rect",
                    "xref": "x",
                    "yref": "paper",
                    "x0": x_min,
                    "y0": 0,
                    "x1": newLSL,
                    "y1": 1,
                    "fillcolor": "rgba(0, 123, 197, 0.1)",
                    "line": {"width": 0},
                    "layer": "below",
                }
            )
            shapes_after.append(
                {
                    "type": "rect",
                    "xref": "x",
                    "yref": "paper",
                    "x0": newUSL,
                    "y0": 0,
                    "x1": x_max,
                    "y1": 1,
                    "fillcolor": "rgba(0, 123, 197, 0.1)",
                    "line": {"width": 0},
                    "layer": "below",
                }
            )

        title_after = (
            f"2. Centered Process vs. Required Specs (Tol: {newToleranceTotal:.{dp}f})"
            if np.isfinite(newToleranceTotal)
            else (
                "2. Centered Process (σ=0)"
                if s == 0
                else "2. Centered Process Distribution"
            )
        )
        fig_after.update_layout(
            **layout_defaults,
            title={"text": title_after, "font": {"size": 12}},
            shapes=shapes_after,
        )

        # Plot 3: Frequency Histogram
        fig_hist = None
        if importedData and len(importedData) >= 2:
            fig_hist = go.Figure()
            fig_hist.add_trace(
                go.Histogram(
                    x=importedData,
                    name="Data Count",
                    marker={
                        "color": "rgba(0, 123, 197, 0.7)",
                        "line": {"color": "rgba(0, 70, 130, 0.8)", "width": 0.5},
                    },
                )
            )

            shapes_hist = [
                {
                    "type": "line",
                    "x0": x_bar,
                    "x1": x_bar,
                    "y0": 0,
                    "y1": 1,
                    "yref": "paper",
                    "line": {"color": "#DC2626", "width": 1.5, "dash": "dash"},
                },
                {
                    "type": "line",
                    "x0": LSL,
                    "x1": LSL,
                    "y0": 0,
                    "y1": 1,
                    "yref": "paper",
                    "line": {"color": "#059669", "width": 1.5, "dash": "dot"},
                },
                {
                    "type": "line",
                    "x0": USL,
                    "x1": USL,
                    "y0": 0,
                    "y1": 1,
                    "yref": "paper",
                    "line": {"color": "#059669", "width": 1.5, "dash": "dot"},
                },
            ]
            annotations_hist = [
                {
                    "x": x_bar,
                    "y": 1.02,
                    "yref": "paper",
                    "text": "Mean",
                    "showarrow": False,
                    "font": {"size": 10, "color": "#DC2626"},
                },
                {
                    "x": LSL,
                    "y": 1.02,
                    "yref": "paper",
                    "text": "LSL",
                    "showarrow": False,
                    "font": {"size": 10, "color": "#059669"},
                    "xanchor": "right",
                },
                {
                    "x": USL,
                    "y": 1.02,
                    "yref": "paper",
                    "text": "USL",
                    "showarrow": False,
                    "font": {"size": 10, "color": "#059669"},
                    "xanchor": "left",
                },
            ]

            fig_hist.update_layout(
                title="Data Frequency Distribution",
                xaxis={"title": "Value", "range": [x_min, x_max], "zeroline": False},
                yaxis={"title": "Frequency (Count)", "fixedrange": True},
                bargap=0.05,
                shapes=shapes_hist,
                annotations=annotations_hist,
                margin={"t": 50, "b": 50, "l": 50, "r": 20},
            )

        return fig_before, fig_after, fig_hist


# --- Summary Panel Logic ---
# Ported from 'updateSummaryPanel'
def get_summary_panel_content(results):
    shiftValue = results.get("shiftValue", np.nan)
    s = results.get("s", np.nan)
    CpkCurrent = results.get("CpkCurrent", np.nan)
    target_index_value = results.get("target_index_value", np.nan)
    LSL, USL = results.get("lsl", np.nan), results.get("usl", np.nan)
    minus3s, plus3s = results.get("minus3s", np.nan), results.get("plus3s", np.nan)
    minus4s, plus4s = results.get("minus4s", np.nan), results.get("plus4s", np.nan)
    newToleranceTotal = results.get("newToleranceTotal", np.nan)
    T_drawing = results.get("T_drawing", np.nan)
    dp = results.get("dp", 3)
    hypothesisResult = results.get("hypothesisResult", {})

    calculation_invalid = (
        not np.isfinite(shiftValue)
        or (not np.isfinite(CpkCurrent) and s != 0)
        or not np.isfinite(newToleranceTotal)
        or not np.isfinite(s)
        or s < 0
    )

    if calculation_invalid:
        return {
            "verdict": "INVALID INPUTS",
            "verdict_color": "red",
            "centering": "Calculation failed due to invalid or incomplete inputs (e.g., negative Std Dev, NaN values).",
            "capability": "",
            "robustness": "",
            "robustness_class": "status-red",
            "tolerance": "",
            "hypothesis": "",
            "recommendations": [
                "<li>Enter valid numeric inputs to see recommendations. Ensure Standard Deviation is not negative.</li>"
            ],
        }

    recommendations = []
    is_good = True
    is_marginal = False

    # Centering
    if s == 0:
        centering_text = f'<span style="color: green; font-weight: bold;">Excellent:</span> Process has zero variation and is centered{f" (but requires shift of {shiftValue:.{dp}f})." if shiftValue != 0 else "."}'
        if shiftValue != 0:
            recommendations.append(
                f"Adjust process mean by <b>{shiftValue:.{dp}f}</b> to align with T<sub>m</sub>."
            )
            is_marginal = True
    elif abs(shiftValue) < (s * 0.05):
        centering_text = '<span style="color: green; font-weight: bold;">Excellent:</span> Process is well-centered.'
    else:
        centering_text = f'<span style="color: orange; font-weight: bold;">Needs Adjustment:</span> Mean is off-target by <b>{shiftValue:.{dp}f}</b>. Adjustment of <b>{abs(shiftValue):.{dp}f} {"UP (+)" if shiftValue < 0 else "DOWN (-)"}</b> is required.'
        recommendations.append(
            f"Adjust process mean by <b>{shiftValue:.{dp}f}</b> to align with T<sub>m</sub>."
        )
        is_marginal = True

    # Capability
    if s == 0:
        capability_text = '<span style="color: green; font-weight: bold;">Perfect Capability (σ=0):</span> Index is effectively infinite (∞).'
    elif np.isfinite(CpkCurrent) and CpkCurrent >= target_index_value:
        capability_text = f'<span style="color: green; font-weight: bold;">Capable:</span> Current index of <b>{CpkCurrent:.{dp}f}</b> meets target of <b>{target_index_value:.2f}</b>.'
    elif np.isfinite(CpkCurrent) and CpkCurrent >= 1.33:
        capability_text = f'<span style="color: orange; font-weight: bold;">Marginally Capable:</span> Index of <b>{CpkCurrent:.{dp}f}</b> is acceptable but below target ({target_index_value:.2f}).'
        recommendations.append(
            "Improve stability or reduce variation (σ) to meet capability target."
        )
        is_marginal = True
    else:
        cpk_display = f"{CpkCurrent:.{dp}f}" if np.isfinite(CpkCurrent) else "N/A"
        capability_text = f'<span style="color: red; font-weight: bold;">Not Capable:</span> Index of <b>{cpk_display}</b> is below target ({target_index_value:.2f}). High risk of defects.'
        recommendations.append(
            "Urgent action required to reduce variation (σ) and/or re-center mean."
        )
        is_good = False

    # Robustness
    robustness_text = ""
    robustness_class = ""
    if s == 0:
        robustness_text = "ROBUST: Process has zero variation."
        robustness_class = "status-green"
    elif all(np.isfinite(v) for v in [LSL, USL, minus3s, plus3s, minus4s, plus4s]):
        if LSL < minus3s and USL > plus3s:
            if LSL < minus4s and USL > plus4s:
                robustness_text = "ROBUST: The ±4σ process spread is contained within specification limits."
                robustness_class = "status-green"
            else:
                robustness_text = "MARGINAL: The ±3σ spread is contained, but ±4σ is NOT. Low tolerance for future shifts."
                robustness_class = "status-yellow"
        else:
            robustness_text = (
                "NOT ROBUST: The ±3σ process spread breaches the specification limits."
            )
            robustness_class = "status-red"
    else:
        robustness_text = "Robustness check skipped due to invalid limits/spread."

    # Tolerance
    if s == 0:
        tolerance_text = '<span style="color: green; font-weight: bold;">Adequate:</span> Zero variation requires zero tolerance.'
    elif np.isfinite(T_drawing) and newToleranceTotal <= T_drawing:
        tolerance_text = f'<span style="color: green; font-weight: bold;">Adequate:</span> Current tolerance of <b>{T_drawing:.{dp}f}</b> is sufficient.'
    elif np.isfinite(T_drawing):
        tolerance_text = f'<span style="color: red; font-weight: bold;">Inadequate:</span> Tolerance is too tight. Requires minimum of <b>{newToleranceTotal:.{dp}f}</b>.'
        recommendations.append(
            "Widen specification range or fundamentally reduce process variation (σ)."
        )
        is_good = False
    else:
        tolerance_text = "Tolerance check skipped due to invalid limits."

    # Hypothesis Test
    p_value = hypothesisResult.get("p_value", np.nan)
    alpha = hypothesisResult.get("alpha", np.nan)
    z_stat = hypothesisResult.get("z_stat", np.nan)

    if np.isfinite(p_value) and np.isfinite(alpha) and np.isfinite(z_stat):
        if p_value < alpha:
            hypothesis_text = f'<span style="color: orange; font-weight: bold;">Reject H₀:</span> With a p-value of <b>{p_value:.3e}</b> (which is < α={alpha:.2f}), there is significant evidence that the process mean has shifted from the target.'
            is_marginal = True
        else:
            hypothesis_text = f'<span style="color: green; font-weight: bold;">Fail to Reject H₀:</span> With a p-value of <b>{p_value:.3e}</b> (which is >= α={alpha:.2f}), there is no significant evidence that the mean has shifted from the target.'
        hypothesis_text += f" (Z-statistic: {z_stat:.3f})"
    else:
        hypothesis_text = "Hypothesis test skipped (requires n>=2 and valid inputs)."

    # Final Verdict
    if not is_good:
        verdict_text = "ACTION REQUIRED"
        verdict_color = "red"
    elif is_marginal:
        verdict_text = "MARGINAL"
        verdict_color = "orange"
    else:
        verdict_text = "PROCESS HEALTH: GOOD"
        verdict_color = "green"

    if not recommendations:
        recommendations.append(
            "Process appears to meet target criteria based on current data. Monitor for stability."
        )

    return {
        "verdict": verdict_text,
        "verdict_color": verdict_color,
        "centering": centering_text,
        "capability": capability_text,
        "robustness": robustness_text,
        "robustness_class": robustness_class,  # This part is tricky to style in st.markdown
        "tolerance": tolerance_text,
        "hypothesis": hypothesis_text,
        "recommendations": [f"<li>{r}</li>" for r in recommendations],
    }


# --- Export Logic ---
# Ported from 'exportManager'
class ExportManager:
    def __init__(self):
        self.styles = {
            "title": {
                "font": Font(bold=True, sz=16, color="1F2937"),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
            "subtitle": {
                "font": Font(sz=10, color="6B7280"),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
            "header": {
                "font": Font(bold=True, color="FFFFFFFF"),
                "fill": PatternFill(
                    start_color="4B5563", end_color="4B5563", fill_type="solid"
                ),
                "alignment": Alignment(
                    horizontal="center", vertical="center", wrapText=True
                ),
                "border": Border(
                    bottom=Side(style="thin"),
                    top=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                ),
            },
            "subheader": {
                "font": Font(bold=True, color="111827", sz=12),
                "fill": PatternFill(
                    start_color="E5E7EB", end_color="E5E7EB", fill_type="solid"
                ),
                "alignment": Alignment(horizontal="center", vertical="center"),
                "border": Border(bottom=Side(style="medium", color="007BC5")),
            },
            "metricLabel": {
                "font": Font(bold=True),
                "alignment": Alignment(horizontal="right", vertical="center"),
            },
            "good": {
                "fill": PatternFill(
                    start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"
                ),
                "font": Font(color="065F46", bold=True),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
            "marginal": {
                "fill": PatternFill(
                    start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"
                ),
                "font": Font(color="92400E", bold=True),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
            "bad": {
                "fill": PatternFill(
                    start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"
                ),
                "font": Font(color="991B1B", bold=True),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
            "dataCell": {
                "border": Border(
                    bottom=Side(style="dotted", color="D1D5DB"),
                    top=Side(style="dotted", color="D1D5DB"),
                    left=Side(style="dotted", color="D1D5DB"),
                    right=Side(style="dotted", color="D1D5DB"),
                ),
                "alignment": Alignment(vertical="center"),
            },
            "wrap": {"alignment": Alignment(wrapText=True, vertical="top")},
            "infinity": {
                "font": Font(sz=14),
                "alignment": Alignment(horizontal="right", vertical="center"),
            },
        }
        self.number_formats = {
            "integer": "0",
            "ppm": "#,##0",
            "scientific": "0.00E+00",
            "dateTime": "yyyy-mm-dd hh:mm:ss",
        }

    def _get_num_style(self, dp=3):
        dp = int(dp) if dp is not None else 3
        return {
            "number_format": f"0.{'0' * dp}",
            "alignment": Alignment(horizontal="right", vertical="center"),
        }

    def _get_perc_style(self, dp=3):
        dp = dp if dp is not None and isinstance(dp, int) else 3
        return {
            "number_format": f"0.{'0' * dp}%",
            "alignment": Alignment(horizontal="right", vertical="center"),
        }

    def _apply_styles(self, ws, data_with_styles):
        max_col_width = {}
        for r_idx, row in enumerate(data_with_styles, 1):
            for c_idx, cell_data in enumerate(row, 1):
                if not cell_data:
                    continue

                cell = ws.cell(row=r_idx, column=c_idx, value=cell_data["v"])
                cell.style = "Normal"  # Reset style

                style_dict = {**self.styles.get("dataCell", {})}

                if "s" in cell_data:
                    style_dict.update(cell_data["s"])

                if "font" in style_dict:
                    cell.font = style_dict["font"]
                if "fill" in style_dict:
                    cell.fill = style_dict["fill"]
                if "alignment" in style_dict:
                    cell.alignment = style_dict["alignment"]
                if "border" in style_dict:
                    cell.border = style_dict["border"]
                if "number_format" in style_dict:
                    cell.number_format = style_dict["number_format"]

                # Auto-fit columns
                cell_len = len(str(cell.value))
                if c_idx not in max_col_width or cell_len > max_col_width[c_idx]:
                    max_col_width[c_idx] = cell_len

        for c_idx, width in max_col_width.items():
            ws.column_dimensions[get_column_letter(c_idx)].width = min(
                max(width + 2, 10), 60
            )

    def _create_cell(self, value, style_keys=None, extra_styles=None):
        if style_keys is None:
            style_keys = []
        if extra_styles is None:
            extra_styles = {}

        final_style = {}
        for key in style_keys:
            if key in self.styles:
                final_style.update(self.styles[key])
        final_style.update(extra_styles)

        display_value = value
        if isinstance(value, (int, float)):
            if not np.isfinite(value):
                display_value = "∞" if value > 0 else "-∞"
                final_style.update(self.styles.get("infinity", {}))
        elif value is None:
            display_value = ""

        return {"v": display_value, "s": final_style}

    def export_current_results(self, results, summary):
        dp = results.get("dp", 3)
        dp = int(dp) if dp is not None else 3  # Ensure dp is an integer
        num_style = self._get_num_style(dp)
        num_style_s = self._get_num_style(dp + 2)
        num_style_pval = {
            "number_format": self.number_formats["scientific"],
            "alignment": self._get_num_style(dp)["alignment"],
        }
        ppm_style = {
            "number_format": self.number_formats["ppm"],
            "alignment": self._get_num_style(dp)["alignment"],
        }
        int_style = {
            "number_format": self.number_formats["integer"],
            "alignment": self._get_num_style(dp)["alignment"],
        }
        perc_style = self._get_perc_style(3)

        verdict_style_key = (
            "bad"
            if summary["verdict_color"] == "red"
            else (
                "marginal"
                if summary["verdict_color"] == "orange"
                else ("good" if summary["verdict_color"] == "green" else "dataCell")
            )
        )
        verdict_style = self.styles.get(verdict_style_key, {})

        cpk_meets_target = (
            np.isfinite(results.get("CpkCurrent", np.nan))
            and np.isfinite(results.get("target_index_value", np.nan))
            and results["CpkCurrent"] >= results["target_index_value"]
        )
        cpk_style = (
            self.styles["good"]
            if results.get("s") == 0 and results.get("CpkCurrent", 0) > 0
            else (self.styles["good"] if cpk_meets_target else self.styles["bad"])
        )

        shift_significant = (
            np.isfinite(results.get("shiftValue", np.nan))
            and results.get("s", 0) > 0
            and abs(results["shiftValue"]) >= (results["s"] * 0.05)
        )
        shift_style = {
            **num_style,
            **(self.styles["marginal"] if shift_significant else {}),
        }

        ppm_below_style = {
            **ppm_style,
            **(self.styles["bad"] if results.get("ppm_below", 0) > 0 else {}),
        }
        ppm_above_style = {
            **ppm_style,
            **(self.styles["bad"] if results.get("ppm_above", 0) > 0 else {}),
        }

        hypo = results.get("hypothesisResult", {})
        hypo_conclusion = "N/A"
        hypo_style = self.styles["wrap"]
        if np.isfinite(hypo.get("p_value", np.nan)) and np.isfinite(
            hypo.get("alpha", np.nan)
        ):
            hypo_conclusion = (
                "Reject Null Hypothesis (Significant Shift)"
                if hypo["p_value"] < hypo["alpha"]
                else "Fail to Reject Null Hypothesis (No Significant Shift)"
            )
        elif results.get("s") == 0:
            hypo_conclusion = (
                "Reject Null Hypothesis (Significant Shift)"
                if results.get("shiftValue") != 0
                else "Fail to Reject Null Hypothesis (No Shift)"
            )

        if "Reject" in hypo_conclusion:
            hypo_style.update(self.styles["marginal"])
        elif "Fail" in hypo_conclusion:
            hypo_style.update(self.styles["good"])

        data = [
            [
                self._create_cell(
                    f"Capability Analysis Report: {results.get('measurement_name', 'Unnamed')}",
                    ["title"],
                ),
                None,
                None,
            ],
            [
                self._create_cell(
                    f"Analysis Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    ["subtitle"],
                ),
                None,
                None,
            ],
            [],
            [
                self._create_cell("OVERALL ASSESSMENT", ["subheader"]),
                self._create_cell(summary["verdict"], [verdict_style_key]),
                None,
            ],
            [],
            [self._create_cell("INPUT PARAMETERS", ["subheader"]), None, None],
            [
                self._create_cell("Metric", ["header"]),
                self._create_cell("Value", ["header"]),
                self._create_cell("Notes", ["header"]),
            ],
            [
                self._create_cell("Measurement Name", ["metricLabel"]),
                self._create_cell(results.get("measurement_name")),
                None,
            ],
            [
                self._create_cell("Target Mean (Tm)", ["metricLabel"]),
                self._create_cell(results.get("tm"), extra_styles=num_style),
                None,
            ],
            [
                self._create_cell("LSL", ["metricLabel"]),
                self._create_cell(results.get("lsl"), extra_styles=num_style),
                None,
            ],
            [
                self._create_cell("USL", ["metricLabel"]),
                self._create_cell(results.get("usl"), extra_styles=num_style),
                None,
            ],
        ]

        if results.get("mode") == "manual":
            data.extend(
                [
                    [
                        self._create_cell("Measured Mean (x̄)", ["metricLabel"]),
                        self._create_cell(results.get("x_bar"), extra_styles=num_style),
                        None,
                    ],
                    [
                        self._create_cell("Std Deviation (s)", ["metricLabel"]),
                        self._create_cell(results.get("s"), extra_styles=num_style_s),
                        None,
                    ],
                    [
                        self._create_cell("Sample Size (n)", ["metricLabel"]),
                        self._create_cell(
                            results.get("n_samples"), extra_styles=int_style
                        ),
                        None,
                    ],
                ]
            )
        else:
            data.extend(
                [
                    [
                        self._create_cell("Data Source", ["metricLabel"]),
                        self._create_cell("Imported Data"),
                        None,
                    ],
                    [
                        self._create_cell("(Calculated) Mean (x̄)", ["metricLabel"]),
                        self._create_cell(results.get("x_bar"), extra_styles=num_style),
                        self._create_cell("From imported data", ["wrap"]),
                    ],
                    [
                        self._create_cell("(Calculated) Std Dev (s)", ["metricLabel"]),
                        self._create_cell(results.get("s"), extra_styles=num_style_s),
                        self._create_cell("From imported data", ["wrap"]),
                    ],
                    [
                        self._create_cell(
                            "(Calculated) Sample Size (n)", ["metricLabel"]
                        ),
                        self._create_cell(
                            results.get("n_samples"), extra_styles=int_style
                        ),
                        self._create_cell("From imported data", ["wrap"]),
                    ],
                ]
            )

        data.extend(
            [
                [
                    self._create_cell("Target Index Type", ["metricLabel"]),
                    self._create_cell(results.get("target_index_type")),
                    None,
                ],
                [
                    self._create_cell("Target Index Value", ["metricLabel"]),
                    self._create_cell(
                        results.get("target_index_value"),
                        extra_styles=self._get_num_style(2),
                    ),
                    None,
                ],
                [
                    self._create_cell("Confidence Level (%)", ["metricLabel"]),
                    self._create_cell(
                        results.get("confidence_level"), extra_styles=int_style
                    ),
                    None,
                ],
                [
                    self._create_cell("Distribution", ["metricLabel"]),
                    self._create_cell(results.get("distribution")),
                    None,
                ],
                [
                    self._create_cell("Hypothesis Type", ["metricLabel"]),
                    self._create_cell(results.get("hypothesis_type")),
                    None,
                ],
                [],
                [self._create_cell("CALCULATED RESULTS", ["subheader"]), None, None],
                [
                    self._create_cell("Metric", ["header"]),
                    self._create_cell("Value", ["header"]),
                    self._create_cell("Notes", ["header"]),
                ],
                [
                    self._create_cell("Shift Required (Tm - x̄)", ["metricLabel"]),
                    self._create_cell(
                        results.get("shiftValue"), extra_styles=shift_style
                    ),
                    None,
                ],
                [
                    self._create_cell("Drawing Tolerance (USL - LSL)", ["metricLabel"]),
                    self._create_cell(results.get("T_drawing"), extra_styles=num_style),
                    None,
                ],
                [
                    self._create_cell("6σ Spread (6 * s)", ["metricLabel"]),
                    self._create_cell(
                        results.get("sixSigmaSpread"), extra_styles=num_style
                    ),
                    None,
                ],
                [
                    self._create_cell("8σ Spread (8 * s)", ["metricLabel"]),
                    self._create_cell(
                        results.get("eightSigmaSpread"), extra_styles=num_style
                    ),
                    None,
                ],
                [
                    self._create_cell("Cp (Potential Capability)", ["metricLabel"]),
                    self._create_cell(results.get("Cp"), extra_styles=num_style),
                    None,
                ],
                [
                    self._create_cell(
                        f"{results.get('target_index_type')} (Actual Capability)",
                        ["metricLabel"],
                    ),
                    self._create_cell(
                        results.get("CpkCurrent"),
                        extra_styles={**num_style, **cpk_style},
                    ),
                    self._create_cell(
                        f"Target: {results.get('target_index_value', 0):.2f}", ["wrap"]
                    ),
                ],
                [
                    self._create_cell("Required Tolerance", ["metricLabel"]),
                    self._create_cell(
                        results.get("newToleranceTotal"), extra_styles=num_style
                    ),
                    self._create_cell(
                        f"For target {results.get('target_index_type')} of {results.get('target_index_value', 0):.2f}",
                        ["wrap"],
                    ),
                ],
                [
                    self._create_cell("Confidence Interval Lower", ["metricLabel"]),
                    self._create_cell(results.get("ci_lower"), extra_styles=num_style),
                    None,
                ],
                [
                    self._create_cell("Confidence Interval Upper", ["metricLabel"]),
                    self._create_cell(results.get("ci_upper"), extra_styles=num_style),
                    None,
                ],
                [],
                [self._create_cell("PROBABILITY & DEFECTS", ["subheader"]), None, None],
                [
                    self._create_cell("Metric", ["header"]),
                    self._create_cell("Value", ["header"]),
                    None,
                ],
                [
                    self._create_cell("Probability < LSL (%)", ["metricLabel"]),
                    self._create_cell(
                        results.get("prob_below"), extra_styles=perc_style
                    ),
                    None,
                ],
                [
                    self._create_cell("Probability > USL (%)", ["metricLabel"]),
                    self._create_cell(
                        results.get("prob_above"), extra_styles=perc_style
                    ),
                    None,
                ],
                [
                    self._create_cell("PPM < LSL", ["metricLabel"]),
                    self._create_cell(
                        results.get("ppm_below"), extra_styles=ppm_below_style
                    ),
                    None,
                ],
                [
                    self._create_cell("PPM > USL", ["metricLabel"]),
                    self._create_cell(
                        results.get("ppm_above"), extra_styles=ppm_above_style
                    ),
                    None,
                ],
                [],
                [
                    self._create_cell(
                        "HYPOTHESIS TEST (Mean vs Target)", ["subheader"]
                    ),
                    None,
                    None,
                ],
                [
                    self._create_cell("Metric", ["header"]),
                    self._create_cell("Value", ["header"]),
                    None,
                ],
                [
                    self._create_cell("Z-Statistic", ["metricLabel"]),
                    self._create_cell(
                        hypo.get("z_stat"), extra_styles=self._get_num_style(4)
                    ),
                    None,
                ],
                [
                    self._create_cell("P-Value", ["metricLabel"]),
                    self._create_cell(hypo.get("p_value"), extra_styles=num_style_pval),
                    None,
                ],
                [
                    self._create_cell("Alpha", ["metricLabel"]),
                    self._create_cell(
                        hypo.get("alpha"), extra_styles=self._get_num_style(2)
                    ),
                    None,
                ],
                [
                    self._create_cell("Conclusion", ["metricLabel"]),
                    self._create_cell(hypo_conclusion, extra_styles=hypo_style),
                    None,
                ],
            ]
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "Capability Report"

        self._apply_styles(ws, data)

        # Apply merges
        ws.merge_cells("A1:C1")
        ws.merge_cells("A2:C2")
        ws.merge_cells("B4:C4")
        ws.merge_cells("A6:C6")
        ws.merge_cells("A24:C24")
        ws.merge_cells("A35:C35")
        ws.merge_cells("A41:C41")

        # Apply row heights
        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 14
        ws.row_dimensions[3].height = 6
        ws.row_dimensions[4].height = 20

        # Save to memory buffer
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def export_selected_history(self, history_data):
        headers = [
            "Timestamp",
            "Measurement_Name",
            "Verdict",
            "Cp",
            "Cpk/Ppk",
            "Target_Index_Type",
            "Target_Index_Value",
            "Required_Shift",
            "Target_Mean_Tm",
            "LSL",
            "USL",
            "Measured_Mean_xbar",
            "Std_Dev_s",
            "Sample_Size_n",
            "PPM_Below_LSL",
            "PPM_Above_USL",
            "Distribution",
            "Confidence_Level",
            "Hypothesis_Type",
            "Z_Stat",
            "P_Value",
            "Alpha",
            "Hypo_Conclusion",
        ]

        data = [[self._create_cell(h, ["header"]) for h in headers]]

        for entry in history_data:
            dp = entry.get("dp", 3)
            dp = int(dp) if dp is not None else 3  # Ensure dp is an integer
            verdict = entry.get("verdict", "N/A")
            verdict_style_key = (
                "bad"
                if "ACTION" in verdict or "INVALID" in verdict
                else (
                    "marginal"
                    if "MARGINAL" in verdict
                    else ("good" if "GOOD" in verdict else "dataCell")
                )
            )

            num_style = self._get_num_style(dp)
            num_style_more = self._get_num_style(dp + 1)
            ppm_style = {
                "number_format": self.number_formats["ppm"],
                "alignment": self._get_num_style(dp)["alignment"],
            }
            int_style = {
                "number_format": self.number_formats["integer"],
                "alignment": self._get_num_style(dp)["alignment"],
            }
            num_style_pval = {
                "number_format": self.number_formats["scientific"],
                "alignment": self._get_num_style(dp)["alignment"],
            }

            cpk_meets = (
                np.isfinite(entry.get("CpkCurrent", np.nan))
                and np.isfinite(entry.get("target_index_value", np.nan))
                and entry["CpkCurrent"] >= entry["target_index_value"]
            )
            cpk_style = (
                self.styles["good"]
                if entry.get("s") == 0 and entry.get("CpkCurrent", 0) > 0
                else (self.styles["good"] if cpk_meets else self.styles["bad"])
            )

            shift_sig = (
                np.isfinite(entry.get("shiftValue", np.nan))
                and entry.get("s", 0) > 0
                and abs(entry["shiftValue"]) >= (entry["s"] * 0.05)
            )
            shift_style = {
                **num_style,
                **(self.styles["marginal"] if shift_sig else {}),
            }

            ppm_below_style = {
                **ppm_style,
                **(self.styles["bad"] if entry.get("ppm_below", 0) > 0 else {}),
            }
            ppm_above_style = {
                **ppm_style,
                **(self.styles["bad"] if entry.get("ppm_above", 0) > 0 else {}),
            }

            hypo = entry.get("hypothesisResult", {})
            hypo_conclusion = ""
            hypo_style_key = "dataCell"
            if np.isfinite(hypo.get("p_value", np.nan)) and np.isfinite(
                hypo.get("alpha", np.nan)
            ):
                hypo_conclusion = (
                    "Reject H0"
                    if hypo["p_value"] < hypo["alpha"]
                    else "Fail to Reject H0"
                )
            elif entry.get("s") == 0:
                hypo_conclusion = (
                    "Reject H0" if entry.get("shiftValue") != 0 else "Fail to Reject H0"
                )

            if "Reject" in hypo_conclusion:
                hypo_style_key = "marginal"
            elif "Fail" in hypo_conclusion:
                hypo_style_key = "good"

            row = [
                self._create_cell(
                    datetime.datetime.fromisoformat(entry.get("id"))
                    if entry.get("id")
                    else None,
                    extra_styles={"number_format": self.number_formats["dateTime"]},
                ),
                self._create_cell(entry.get("measurement_name", "")),
                self._create_cell(verdict, [verdict_style_key]),
                self._create_cell(entry.get("Cp"), extra_styles=num_style),
                self._create_cell(
                    entry.get("CpkCurrent"), extra_styles={**num_style, **cpk_style}
                ),
                self._create_cell(entry.get("target_index_type", "")),
                self._create_cell(
                    entry.get("target_index_value"), extra_styles=self._get_num_style(2)
                ),
                self._create_cell(entry.get("shiftValue"), extra_styles=shift_style),
                self._create_cell(entry.get("tm"), extra_styles=num_style),
                self._create_cell(entry.get("lsl"), extra_styles=num_style),
                self._create_cell(entry.get("usl"), extra_styles=num_style),
                self._create_cell(entry.get("x_bar"), extra_styles=num_style),
                self._create_cell(entry.get("s"), extra_styles=num_style_more),
                self._create_cell(entry.get("n_samples"), extra_styles=int_style),
                self._create_cell(entry.get("ppm_below"), extra_styles=ppm_below_style),
                self._create_cell(entry.get("ppm_above"), extra_styles=ppm_above_style),
                self._create_cell(entry.get("distribution", "")),
                self._create_cell(
                    entry.get("confidence_level"), extra_styles=int_style
                ),
                self._create_cell(entry.get("hypothesis_type", "")),
                self._create_cell(
                    hypo.get("z_stat"), extra_styles=self._get_num_style(4)
                ),
                self._create_cell(hypo.get("p_value"), extra_styles=num_style_pval),
                self._create_cell(
                    hypo.get("alpha"), extra_styles=self._get_num_style(2)
                ),
                self._create_cell(hypo_conclusion, [hypo_style_key]),
            ]
            data.append(row)

        wb = Workbook()
        ws = wb.active
        ws.title = "History Selection"

        self._apply_styles(ws, data)
        ws.row_dimensions[1].height = 20

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer


# --- Sigma Assistant Mascot (Clippy-Style) ---
class SigmaAssistant:
    """Generates a Clippy-style floating Sigma Assistant with CSS animations."""

    # Messages for each state
    STATE_MESSAGES = {
        "idle": "Hello! I'm here to help. Run an analysis to see results!",
        "happy": "Excellent! Process is stable and capable. Great work!",
        "sad": "Action required. The process needs improvement.",
        "thinking": "Analyzing the data... Please wait!",
        "concerned": "The process is marginal. Review recommendations.",
    }

    # State colors (border color for the body)
    STATE_COLORS = {
        "idle": "#3B82F6",  # Blue
        "happy": "#10B981",  # Green
        "sad": "#EF4444",  # Red
        "thinking": "#FBBF24",  # Yellow
        "concerned": "#F97316",  # Orange
    }

    @classmethod
    def render_fixed(cls, state="idle", message=None):
        """
        Render the Clippy-style mascot using st.markdown for TRUE fixed positioning.
        This injects CSS/HTML directly into Streamlit's main page, not an iframe.
        """
        if state not in cls.STATE_MESSAGES:
            state = "idle"

        msg = message if message else cls.STATE_MESSAGES.get(state)
        color = cls.STATE_COLORS.get(state, "#3B82F6")

        # Animation name based on state
        animation_map = {
            "idle": "sigma-bob",
            "happy": "sigma-happy-dance",
            "sad": "sigma-sad-slump",
            "thinking": "sigma-thinking",
            "concerned": "sigma-bob",
        }
        animation = animation_map.get(state, "sigma-bob")

        # Mouth path based on state
        mouth_map = {
            "idle": "M 55 90 Q 70 95 85 90",
            "happy": "M 55 85 C 60 105, 80 105, 85 85",
            "sad": "M 55 95 Q 70 85 85 95",
            "thinking": "M 60 90 L 80 90",
            "concerned": "M 55 93 Q 70 88 85 93",
        }
        mouth = mouth_map.get(state, mouth_map["idle"])

        # Eyebrow transforms based on state
        eyebrow_left = "translate(50, 45)"
        eyebrow_right = "translate(90, 45)"
        if state == "sad":
            eyebrow_left = "translate(45, 43) rotate(15)"
            eyebrow_right = "translate(95, 43) rotate(-15)"
        elif state == "concerned":
            eyebrow_left = "translate(50, 45) rotate(10)"
            eyebrow_right = "translate(90, 45) rotate(-10)"
        elif state == "thinking":
            eyebrow_left = "translate(50, 45) rotate(-10)"
            eyebrow_right = "translate(90, 45) rotate(10)"

        html = f'''
<style>
/* Sigma Assistant Fixed Positioning - injected into Streamlit main page */
.sigma-fixed-container {{
    position: fixed !important;
    bottom: 20px !important;
    right: 20px !important;
    z-index: 999999 !important;
    display: flex;
    flex-direction: column;
    align-items: center;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    pointer-events: none;
}}

.sigma-speech-bubble {{
    background-color: #1F2937;
    color: white;
    padding: 10px 14px;
    border-radius: 10px;
    margin-bottom: 8px;
    max-width: 200px;
    font-size: 0.8rem;
    box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    text-align: center;
    line-height: 1.4;
}}

.sigma-mascot {{
    width: 100px;
    height: 100px;
    cursor: pointer;
    pointer-events: auto;
}}

.sigma-mascot svg {{
    filter: drop-shadow(2px 3px 3px rgba(0, 0, 0, 0.2));
}}

/* Animations */
@keyframes sigma-bob {{
    0%, 100% {{ transform: translateY(0); }}
    50% {{ transform: translateY(-3px); }}
}}

@keyframes sigma-happy-dance {{
    0%, 100% {{ transform: translateY(0) rotate(0); }}
    15% {{ transform: translateY(-5px) rotate(3deg); }}
    30% {{ transform: translateY(0) rotate(0); }}
    45% {{ transform: translateY(-5px) rotate(-3deg); }}
    60% {{ transform: translateY(0) rotate(0); }}
}}

@keyframes sigma-sad-slump {{
    0%, 100% {{ transform: translateY(0) rotate(0); }}
    50% {{ transform: translateY(3px) rotate(-1deg) scaleY(0.96); }}
}}

@keyframes sigma-thinking {{
    0%, 100% {{ transform: rotate(0deg); }}
    25% {{ transform: rotate(1deg); }}
    75% {{ transform: rotate(-1deg); }}
}}

.sigma-animate {{
    animation: {animation} 2s ease-in-out infinite;
}}
</style>

<div class="sigma-fixed-container">
    <div class="sigma-speech-bubble">{msg}</div>
    <div class="sigma-mascot sigma-animate">
        <svg viewBox="-20 -30 150 150" xmlns="http://www.w3.org/2000/svg">
            <defs>
                <filter id="sigma-shadow"><feGaussianBlur in="SourceAlpha" stdDeviation="2"/></filter>
            </defs>
            <!-- Shadow -->
            <ellipse cx="55" cy="115" rx="35" ry="8" fill="black" opacity="0.15" filter="url(#sigma-shadow)"/>
            <!-- Body -->
            <path d="M 20 110 C 20 110, 5 20, 55 20 C 105 20, 90 110, 90 110 Z" 
                  fill="#F9FAFB" stroke="{color}" stroke-width="4" stroke-linejoin="round"/>
            <!-- Face -->
            <g transform="translate(0, -10)">
                <!-- Eyebrows -->
                <g transform="{eyebrow_left}">
                    <path d="M -8 0 Q 0 -5 8 0" fill="none" stroke="#4B5563" stroke-width="3" stroke-linecap="round"/>
                </g>
                <g transform="{eyebrow_right}">
                    <path d="M -8 0 Q 0 -5 8 0" fill="none" stroke="#4B5563" stroke-width="3" stroke-linecap="round"/>
                </g>
                <!-- Eyes -->
                <g transform="translate(50, 60)">
                    <ellipse rx="10" ry="8" fill="white" stroke="#1F2937" stroke-width="1.5"/>
                    <circle cx="1" cy="1" r="4" fill="#1F2937"/>
                    <circle cx="3" cy="-1" r="1.5" fill="white"/>
                </g>
                <g transform="translate(90, 60)">
                    <ellipse rx="10" ry="8" fill="white" stroke="#1F2937" stroke-width="1.5"/>
                    <circle cx="1" cy="1" r="4" fill="#1F2937"/>
                    <circle cx="3" cy="-1" r="1.5" fill="white"/>
                </g>
                <!-- Mouth -->
                <path d="{mouth}" fill="none" stroke="#1E3A8A" stroke-width="2.5" stroke-linecap="round"/>
            </g>
        </svg>
    </div>
</div>
'''
        return html


# --- Chatbot Logic ---
# Ported from 'sigmaAssistant'
class Chatbot:
    def __init__(self):
        # Prepare reference content
        self.reference_content_sections = self._prepare_reference_content()
        self.common_words = set(
            [
                "a",
                "an",
                "the",
                "is",
                "are",
                "what",
                "how",
                "when",
                "where",
                "for",
                "to",
                "of",
                "in",
                "and",
                "or",
                "do",
                "does",
                "can",
                "explain",
                "tell",
                "me",
                "about",
            ]
        )

    def _prepare_reference_content(self):
        # This data is manually extracted from the 'tab-content-reference' HTML
        raw_sections = [
            {
                "context": "Application Context & Usage Guide",
                "text": "This tool is primarily utilized in Six Sigma and Statistical Process Control (SPC) environments for Process Centering and Tolerance Verification.",
            },
            {
                "context": "Application Context & Usage Guide",
                "text": "Quantify Process Drift: Calculate the exact Required Shift (Δ) needed to move the measured process mean (x̄) back to the engineering target (Tm). This quantifies the resultant dimensional change or process deformation that has occurred (e.g., due to tool wear or assembly stress).",
            },
            {
                "context": "Application Context & Usage Guide",
                "text": "Predict Initial State: By using the current measured output (x̄) and the required shift (Δ), one can infer the required initial dimension/setting to achieve the target (Tm) after the process variables have exerted their influence.",
            },
            {
                "context": "Application Context & Usage Guide",
                "text": "Verify Tolerance Adequacy: Determine the minimum Required Tolerance (USL - LSL) necessary for the existing process variation (σ) to meet a desired capability index (Cpk, Ppk).",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "Six Sigma (6σ) is a methodology focused on reducing process variation and improving quality to near perfection.",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "3σ Sigma Level has 2,700 PPM (Defects).",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "4σ Sigma Level has 63 PPM (Defects).",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "5σ Sigma Level has 0.57 PPM (Defects).",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "6σ Sigma Level has 0.002 PPM (Defects).",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "Capability indices quantify how well the natural process variation fits within the required specification limits. The choice of index depends on the time horizon of the data collected.",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "Cp/Cpk/Cmk (Short-Term Capability): These indices use within-subgroup variation, typically reflecting the immediate, inherent capability of the process over a short period, free from common causes of variation.",
            },
            {
                "context": "Concepts Overview: Six Sigma & Capability Indices",
                "text": "Pp/Ppk (Long-Term Performance): These indices use overall process variation, including all potential sources of variation over an extended period. Ppk is always less than or equal to Cpk.",
            },
            {
                "context": "Choosing a Distribution: Normal vs. Lognormal",
                "text": "The Normal (or Gaussian) distribution is symmetric and bell-shaped. It's the most common distribution in quality control. Use When: The process data is symmetrically distributed around the mean. Measurements can theoretically be positive or negative (e.g., positional error). You are analyzing dimensional characteristics of manufactured parts.",
            },
            {
                "context": "Choosing a Distribution: Normal vs. Lognormal",
                "text": "The Lognormal distribution is skewed to the right and is bounded by zero on the left. Use When: The data cannot be negative (e.g., time to failure, concentration levels). The logarithm of the data follows a normal distribution. The process produces many values near the lower bound and a few much larger values.",
            },
            {
                "context": "Additional Metric Definitions",
                "text": "x̄ ± 4σ Spread: This range contains approximately 99.9937% of the process output, providing a more conservative view of the process spread compared to the 6σ range (99.73%).",
            },
            {
                "context": "Additional Metric Definitions",
                "text": "P(x < LSL) & P(x > USL): These values represent the calculated probability that a single part will fall outside the lower (LSL) or upper (USL) specification limits. They are the direct drivers of the PPM defect rate.",
            },
            {
                "context": "Additional Metric Definitions",
                "text": "P(x < Tm): This is the probability that a single measurement will fall below the Target Mean (Tm). It helps assess process bias. In a perfectly centered process, this value would be exactly 50%.",
            },
            {
                "context": "Additional Metric Definitions",
                "text": "PPM (Parts Per Million): PPM is a measure of the defect rate, calculated from the probabilities. For example, a P(x > USL) of 0.001% is equivalent to 10 PPM.",
            },
            {
                "context": "Hypothesis Testing & Confidence Level",
                "text": "This tool performs a Z-test to determine if the true process mean (μ) is statistically different from the engineering Target Mean (Tm), based on your sample data.",
            },
            {
                "context": "Hypothesis Testing & Confidence Level",
                "text": "Null Hypothesis (H₀): μ = Tm. This assumes the true process mean is exactly on target.",
            },
            {
                "context": "Hypothesis Testing & Confidence Level",
                "text": "p-value: The probability of observing your sample results if the null hypothesis were true. A small p-value (typically < 0.05) provides evidence against H₀.",
            },
            {
                "context": "Hypothesis Testing & Confidence Level",
                "text": "Result: Fail to Reject H₀. The p-value is high. There is no statistically significant evidence that the process mean is different from the target. The observed off-center condition could be due to random chance.",
            },
            {
                "context": "Hypothesis Testing & Confidence Level",
                "text": "Result: Reject H₀. The p-value is low. There is statistically significant evidence that the process mean has shifted away from the target. The 'Required Shift (Δ)' is likely a real effect.",
            },
            {
                "context": "When to Use Each Hypothesis Type",
                "text": "Two-Sided (Default): Use when deviations in both directions from the target are of concern. Example: A part's diameter must be exactly Tm.",
            },
            {
                "context": "When to Use Each Hypothesis Type",
                "text": "Upper-Sided: Use when it is critical to assess if the mean is significantly above the target. Example: The level of a contaminant must not exceed Tm.",
            },
            {
                "context": "When to Use Each Hypothesis Type",
                "text": "Lower-Sided: Use when it is critical to assess if the mean is significantly below the target. Example: Breaking strength must be at least Tm.",
            },
        ]

        return [
            {"original": s["text"], "lower": s["text"].lower(), "context": s["context"]}
            for s in raw_sections
            if len(s["text"]) > 10
        ]

    def get_response(self, user_input):
        user_input_lower = user_input.lower()
        if not user_input_lower:
            return None

        # Simple keyword extraction
        keywords = [
            word
            for word in re.split(r"[\s,?\-.();:]+", user_input_lower)
            if word and len(word) > 2 and word not in self.common_words
        ]

        if not keywords:
            return "Please ask a more specific question using keywords like 'Cp', 'Lognormal', 'Hypothesis', etc."

        best_match = None
        highest_score = 0

        for section in self.reference_content_sections:
            current_score = 0
            for keyword in keywords:
                if keyword in section["lower"]:
                    current_score += 1
                    # Bonus for matching context
                    if keyword in section["context"].lower():
                        current_score += 1

            if current_score > highest_score:
                highest_score = current_score
                best_match = section
            elif (
                current_score == highest_score
                and best_match
                and len(section["original"]) < len(best_match["original"])
            ):
                best_match = section  # Prefer shorter, more direct answers

        if best_match:
            context_prefix = f'Regarding "{best_match["context"]}": '
            answer = best_match["original"]
            truncated_answer = (answer[:247] + "...") if len(answer) > 250 else answer
            return f"{context_prefix} {truncated_answer}"
        else:
            return "Sorry, I couldn't find that specific topic in the reference guide. Try keywords like 'Cp', 'Cpk', 'Six Sigma', 'hypothesis', 'distribution', 'LSL', 'tolerance', etc."


# --- Main App ---

# Initialize calculators and managers
calc = StatisticalCalculator()
plotter = PlotManager()
exporter = ExportManager()
bot = Chatbot()

# Set page configuration
st.set_page_config(
    page_title="Process Capability Analyzer",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={"Get Help": None, "Report a bug": None, "About": None},
)

# Hide Streamlit's default Deploy button and menu, reduce top padding
st.markdown(
    """
<style>
    /* Hide Deploy button */
    .stDeployButton {display: none !important;}
    /* Hide hamburger menu */
    #MainMenu {display: none !important;}
    /* Hide header completely and remove space */
    header {display: none !important;}
    /* Hide footer */
    footer {display: none !important;}
    /* Reduce top padding of main content */
    .stMainBlockContainer {padding-top: 1rem !important;}
    .block-container {padding-top: 1rem !important;}
    /* Remove any extra top margin in the app view */
    .stAppViewContainer {margin-top: -2rem;}
</style>
""",
    unsafe_allow_html=True,
)


# --- Session State Initialization ---
def init_session_state(clear_form=False):
    defaults = {
        "tm": 10.00,
        "lsl": 9.90,
        "usl": 10.10,
        "target_index_value": 1.67,
        "target_index_type": "Cpk",
        "confidence_level": 95.0,
        "distribution": "Normal",
        "hypothesis_type": "Two-Sided",
        "x_bar": 10.00,
        "s": 0.015,
        "n_samples": 30,
        "decimal_places": 3,
        "mode": "Enter Manually",
        "measurement_name": "",
        "raw_data": "",
    }

    if "history" not in st.session_state:
        st.session_state.history = []

    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = []

    if "results" not in st.session_state:
        st.session_state.results = {}
        st.session_state.summary = {}
        st.session_state.figs = {}

    # Sigma Assistant mascot state
    if "mascot_state" not in st.session_state:
        st.session_state.mascot_state = "idle"
        st.session_state.mascot_cp = 1.0
        st.session_state.mascot_message = None

    if clear_form:
        st.session_state.results = {}
        st.session_state.summary = {}
        st.session_state.figs = {}
        st.session_state.chat_messages = []
        for key, value in defaults.items():
            st.session_state[key] = value
    else:
        for key, value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value


init_session_state()

# --- Main App UI ---
st.title("Statistical Process Capability & Optimization Tool")

# Define Tabs
tab_analysis, tab_data, tab_viz, tab_history, tab_ref = st.tabs(
    ["Analysis & Report", "Data Worksheet", "Visualization", "History", "Reference"]
)

# --- Tab 1: Analysis & Report ---
with tab_analysis:
    # Display Error Messages
    if st.session_state.results and st.session_state.results.get("error"):
        st.error(f"**Analysis Error:** {st.session_state.results['error']}")

    main_cols = st.columns([1, 1, 1])

    # --- Column 1: Input Parameters ---
    with main_cols[0]:
        st.header("I. Input Parameters")

        with st.form(key="analysis_form"):
            st.markdown(
                """
                <p style="font-size: 0.9rem; font-style: italic; color: #555;">
                Define product <b>specifications</b> and <b>measured process performance</b> data.
                </p>
                """,
                unsafe_allow_html=True,
            )

            with st.container(border=True):
                st.subheader("1. Specifications")
                spec_cols = st.columns(3)
                with spec_cols[0]:
                    st.number_input(
                        "Tₘ (Target Mean)",
                        step=0.01,
                        key="tm",
                        help="The desired, ideal center of your process distribution.",
                    )
                with spec_cols[1]:
                    st.number_input(
                        "LSL (Lower Spec)",
                        step=0.01,
                        key="lsl",
                        help="The minimum acceptable value for your measurement.",
                    )
                with spec_cols[2]:
                    st.number_input(
                        "USL (Upper Spec)",
                        step=0.01,
                        key="usl",
                        help="The maximum acceptable value for your measurement.",
                    )

            with st.container(border=True):
                st.subheader("2. Data & Goals")
                st.radio(
                    "Data Input Mode",
                    ["Enter Manually", "Use Data Worksheet"],
                    key="mode",
                    horizontal=True,
                )

                if st.session_state.mode == "Enter Manually":
                    data_cols = st.columns(2)
                    with data_cols[0]:
                        st.number_input(
                            "x̄ (Measured Mean)",
                            step=0.01,
                            key="x_bar",
                            help="The average value calculated from your sample data.",
                        )
                    with data_cols[1]:
                        st.number_input(
                            "σ (Std Dev)",
                            step=0.001,
                            min_value=0.0,
                            format="%.5f",
                            key="s",
                            help="A measure of the amount of variation or dispersion of a set of values.",
                        )

                goal_cols = st.columns(2)
                with goal_cols[0]:
                    st.number_input(
                        "Target Index",
                        step=0.01,
                        key="target_index_value",
                        help="The minimum capability value (e.g., Cpk 1.67) you aim for your process to achieve.",
                    )
                with goal_cols[1]:
                    st.selectbox(
                        "Index Type",
                        ["Cpk", "Cmk", "Ppk"],
                        key="target_index_type",
                        help="Capability Index Type: Cpk (short-term) or Ppk (long-term).",
                    )

            with st.container(border=True):
                st.subheader("3. Statistical Settings")
                stat_cols_1 = st.columns(4)
                with stat_cols_1[0]:
                    st.number_input(
                        "n (Samples)",
                        step=1,
                        min_value=2,
                        key="n_samples",
                        help="The number of data points in your sample. Must be >= 2.",
                    )
                with stat_cols_1[1]:
                    st.number_input(
                        "CL (%)",
                        min_value=1.0,
                        max_value=99.9,
                        step=0.1,
                        key="confidence_level",
                        help="Confidence Level for the Mean's Confidence Interval. 95% is common.",
                    )
                with stat_cols_1[2]:
                    st.number_input(
                        "Decimals",
                        min_value=1,
                        max_value=6,
                        step=1,
                        key="decimal_places",
                    )
                with stat_cols_1[3]:
                    st.selectbox(
                        "Distribution", ["Normal", "Lognormal"], key="distribution"
                    )

                st.selectbox(
                    "Hypothesis (μ vs Tₘ)",
                    options=["Two-Sided", "Upper-Sided", "Lower-Sided"],
                    format_func=lambda x: (
                        f"{x} (μ ≠ Tₘ)"
                        if x == "Two-Sided"
                        else (
                            f"{x} (μ > Tₘ)" if x == "Upper-Sided" else f"{x} (μ < Tₘ)"
                        )
                    ),
                    key="hypothesis_type",
                )

            # Form submission
            submitted = st.form_submit_button(
                "ANALYZE & PLOT", use_container_width=True, type="primary"
            )

        # Other buttons outside the form
        btn_cols = st.columns(2)
        with btn_cols[0]:
            # Export Button - must be explicit boolean for disabled parameter
            export_disabled = bool(
                not st.session_state.results or st.session_state.results.get("error")
            )
            if not export_disabled:
                try:
                    export_buffer = exporter.export_current_results(
                        st.session_state.results, st.session_state.summary
                    )
                    st.download_button(
                        label="EXPORT XLSX",
                        data=export_buffer,
                        file_name=f"Capability_Report_{st.session_state.measurement_name or 'Current'}_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        disabled=export_disabled,
                    )
                except Exception as e:
                    import traceback

                    traceback.print_exc()  # Print full traceback to terminal
                    st.error(f"Could not generate export: {e}")
                    st.button("EXPORT XLSX", use_container_width=True, disabled=True)
            else:
                st.button(
                    "EXPORT XLSX",
                    use_container_width=True,
                    disabled=True,
                    help="Run a valid analysis to enable export.",
                )

        with btn_cols[1]:
            # Reset Button
            if st.button("RESET", use_container_width=True):
                init_session_state(clear_form=True)
                st.rerun()

    # --- Analysis Logic ---
    if submitted:
        # User clicked Analyze, so we run calculations
        st.session_state.results = calc.calculate(st.session_state)

        if not st.session_state.results.get("error"):
            # Get summary
            st.session_state.summary = get_summary_panel_content(
                st.session_state.results
            )
            st.session_state.results["verdict"] = st.session_state.summary.get(
                "verdict", "N/A"
            )

            # Update Sigma Assistant mascot state based on verdict
            verdict = st.session_state.summary.get("verdict", "")
            cp_value = st.session_state.results.get("Cp", 1.0)
            if "GOOD" in verdict:
                st.session_state.mascot_state = "happy"
                st.session_state.mascot_message = None  # Use default happy message
            elif "MARGINAL" in verdict:
                st.session_state.mascot_state = "concerned"
                st.session_state.mascot_message = None  # Use default concerned message
            elif "ACTION" in verdict or "INVALID" in verdict:
                st.session_state.mascot_state = "sad"
                st.session_state.mascot_message = None  # Use default sad message
            else:
                st.session_state.mascot_state = "idle"
                st.session_state.mascot_message = None
            st.session_state.mascot_cp = cp_value if cp_value and cp_value > 0 else 1.0

            # Save to history
            history_entry = st.session_state.results.copy()
            history_entry["id"] = datetime.datetime.now().isoformat()
            if "importedData" in history_entry:
                del history_entry["importedData"]  # Don't save large data array
            st.session_state.history.insert(0, history_entry)
            st.session_state.history = st.session_state.history[:250]  # Limit history

            # Generate plots
            fig_before, fig_after, fig_hist = plotter.update_plots(
                st.session_state.results
            )
            st.session_state.figs = {
                "before": fig_before,
                "after": fig_after,
                "hist": fig_hist,
            }

        else:
            # Clear previous results if new run has errors
            st.session_state.summary = {}
            st.session_state.figs = {}

        st.rerun()  # Rerun to display the new results

    # --- Column 2: Calculated Results ---
    with main_cols[1]:
        st.header("II. Calculated Results")
        st.markdown(
            """
            <p style="font-size: 0.9rem; font-style: italic; color: #555;">
            Key metrics based on the input data, including capability, spread, and recommended adjustments.
            </p>
            """,
            unsafe_allow_html=True,
        )

        res = st.session_state.results
        dp = res.get("dp", 3)

        def format_num(val, default="N/A", dps=None):
            if dps is None:
                dps = dp
            if val is None or not np.isfinite(val):
                return "∞" if val == np.inf else ("-∞" if val == -np.inf else default)
            return f"{val:.{dps}f}"

        if res and not res.get("error"):
            with st.container(border=True):
                st.markdown("**Current Process Metrics**")
                res_cols_1 = st.columns(2)
                with res_cols_1[0]:
                    st.metric(
                        label="Cₚ (Potential)",
                        value=format_num(res.get("Cp")),
                        help="Process Potential (Cp): Measures how capable the process would be if it were perfectly centered.",
                    )
                with res_cols_1[1]:
                    st.metric(
                        label=f"Current Index ({res.get('target_index_type', 'Cpk')})",
                        value=format_num(res.get("CpkCurrent")),
                        help="Process Capability (Cpk/Ppk): Measures the actual process capability, accounting for how centered it is.",
                    )

                st.markdown(
                    f"**6σ Spread:** **`{format_num(res.get('sixSigmaSpread'))}`**",
                    help="The range that contains approximately 99.73% of your process data (Mean ± 3 standard deviations).",
                )
                st.markdown(
                    f"_(x̄ ± 3σ): [ {format_num(res.get('minus3s'))}, {format_num(res.get('plus3s'))} ]_"
                )

                st.markdown(
                    f"**8σ Spread:** **`{format_num(res.get('eightSigmaSpread'))}`**",
                    help="A wider range containing about 99.9937% of data (Mean ± 4 standard deviations).",
                )
                st.markdown(
                    f"_(x̄ ± 4σ): [ {format_num(res.get('minus4s'))}, {format_num(res.get('plus4s'))} ]_"
                )

            res_cols_2 = st.columns(2)
            with res_cols_2[0]:
                with st.container(border=True):
                    st.metric(
                        label="Required Shift (Δ)",
                        value=format_num(res.get("shiftValue")),
                        help="The exact adjustment needed to move the measured process mean to the target mean (Tm).",
                    )
            with res_cols_2[1]:
                with st.container(border=True):
                    st.metric(
                        label=f"Req. Tolerance (Target {res.get('target_index_type')})",
                        value=format_num(res.get("newToleranceTotal")),
                        help="The minimum specification width (USL - LSL) your process needs to achieve its target capability index, given its current standard deviation.",
                    )

            with st.container(border=True):
                ci_label = f"Mean CI @ {res.get('confidence_level')}% ({res.get('hypothesis_type')})"
                ci_value = f"[{format_num(res.get('ci_lower'))}, {format_num(res.get('ci_upper'))}]"
                st.metric(
                    label=ci_label,
                    value=ci_value,
                    help="Confidence Interval (CI) for the Mean: The range within which the true population mean is likely to fall.",
                )

            if res.get("importedData"):
                with st.container(border=True):
                    st.markdown("**Data Summary (Import)**")
                    data_sum_cols = st.columns(3)
                    with data_sum_cols[0]:
                        st.metric("Mean", format_num(res.get("x_bar")))
                    with data_sum_cols[1]:
                        st.metric(
                            "Min", format_num(min(res.get("importedData", [np.nan])))
                        )
                    with data_sum_cols[2]:
                        st.metric(
                            "Max", format_num(max(res.get("importedData", [np.nan])))
                        )

                    if res.get("distribution") == "Lognormal" and np.isfinite(
                        res.get("mu_log", np.nan)
                    ):
                        log_cols = st.columns(2)
                        with log_cols[0]:
                            st.metric("Log-Mean (μ')", format_num(res.get("mu_log")))
                        with log_cols[1]:
                            st.metric(
                                "Log-Std Dev (σ')", format_num(res.get("sigma_log"))
                            )

            with st.container(border=True):
                st.markdown("**Probability & Defect Analysis**")
                prob_cols_1 = st.columns(2)
                with prob_cols_1[0]:
                    st.metric(
                        "P(x < LSL)",
                        f"{res.get('prob_below', 0) * 100:.3f}%",
                        help="The calculated chance that a single part will be produced below the Lower Specification Limit.",
                    )
                    st.metric(
                        "PPM < LSL",
                        f"{res.get('ppm_below', 0):,.0f}",
                        help="The expected number of defective parts per million that will fall below the Lower Specification Limit.",
                    )
                with prob_cols_1[1]:
                    st.metric(
                        "P(x > USL)",
                        f"{res.get('prob_above', 0) * 100:.3f}%",
                        help="The calculated chance that a single part will be produced above the Upper Specification Limit.",
                    )
                    st.metric(
                        "PPM > USL",
                        f"{res.get('ppm_above', 0):,.0f}",
                        help="The expected number of defective parts per million that will fall above the Upper Specification Limit.",
                    )
                st.metric(
                    "P(x < Tₘ)",
                    f"{res.get('prob_below_target', 0) * 100:.1f}%",
                    help="The probability that a single measurement will fall below the Target Mean. If your process is centered on the target, this should be 50%.",
                )
        else:
            st.info("Run analysis to see calculated results.")

    # --- Column 3: Summary & Interpretation ---
    with main_cols[2]:
        st.header("III. Summary & Interpretation")

        summary = st.session_state.summary

        if summary:
            verdict = summary.get("verdict", "ASSESSMENT PENDING")
            color = summary.get("verdict_color", "grey")
            st.markdown(
                f"""
                <div style="padding: 1rem; border-radius: 0.5rem; text-align: center; font-weight: 800; color: white; background-color: {color}; font-size: 1.25rem;">
                    {verdict}
                </div>
                """,
                unsafe_allow_html=True,
            )

            with st.container(border=True):
                st.markdown("<b>1. Process Centering:</b>", unsafe_allow_html=True)
                st.markdown(summary.get("centering", "..."), unsafe_allow_html=True)

                st.markdown(
                    "<br><b>2. Process Capability & Robustness:</b>",
                    unsafe_allow_html=True,
                )
                st.markdown(summary.get("capability", "..."), unsafe_allow_html=True)
                st.markdown(
                    f"<b>{summary.get('robustness', '...')}</b>", unsafe_allow_html=True
                )  # Style is harder here

                st.markdown("<br><b>3. Tolerance Adequacy:</b>", unsafe_allow_html=True)
                st.markdown(summary.get("tolerance", "..."), unsafe_allow_html=True)

                st.markdown(
                    "<br><b>4. Hypothesis Test (μ vs Tₘ):</b>", unsafe_allow_html=True
                )
                st.markdown(summary.get("hypothesis", "..."), unsafe_allow_html=True)

                st.divider()
                st.markdown("<b>Recommendations:</b>", unsafe_allow_html=True)
                st.markdown(
                    f"<ul>{''.join(summary.get('recommendations', []))}</ul>",
                    unsafe_allow_html=True,
                )
        else:
            st.info("Run analysis to see the summary and recommendations.")

# --- Tab 2: Data Worksheet ---
with tab_data:
    st.header("Data Worksheet")
    st.text_input(
        "Measurement Name",
        key="measurement_name",
        placeholder="e.g., Diameter_Part_A",
        help="A descriptive name for your analysis (e.g., 'Diameter_Part_A'). This will be saved in the history.",
    )
    st.text_area(
        "Paste Raw Data (comma, space, or newline separated)",
        key="raw_data",
        height=500,
        placeholder="e.g., 9.98, 10.01, 9.99, ...",
        help="This data will be used when you select 'Use Data Worksheet' on the Analysis tab.",
    )
    if st.session_state.raw_data:
        data_points = calc.parse_raw_data(st.session_state.raw_data)
        st.info(f"Detected **{len(data_points)}** valid numeric data points.")
        if len(data_points) < 2:
            st.warning("At least 2 data points are required for calculation.")


# --- Tab 3: Visualization ---
with tab_viz:
    st.header("Capability Visualization")

    figs = st.session_state.figs

    if figs and figs.get("before") and figs.get("after"):
        viz_cols = st.columns(2)
        with viz_cols[0]:
            st.plotly_chart(figs["before"], use_container_width=True)
        with viz_cols[1]:
            st.plotly_chart(figs["after"], use_container_width=True)

        if figs.get("hist"):
            st.header("Data Distribution Analysis (from Import)")
            st.plotly_chart(figs["hist"], use_container_width=True)
    else:
        st.info(
            "Run analysis on the 'Analysis & Report' tab to generate visualizations."
        )

# --- Tab 4: History ---
with tab_history:
    st.header("Analysis History (Last 250 Runs)")
    st.caption('History is logged only when you click the "ANALYZE & PLOT" button.')

    if not st.session_state.history:
        st.info("No history available. Run an analysis to log it here.")
    else:
        # Filters
        hist_cols = st.columns([2, 1, 1])
        with hist_cols[0]:
            filter_name = st.text_input("Filter by Name")
        with hist_cols[1]:
            filter_verdict = st.selectbox(
                "Filter by Verdict",
                [
                    "all",
                    "PROCESS HEALTH: GOOD",
                    "MARGINAL",
                    "ACTION REQUIRED",
                    "INVALID INPUTS",
                ],
            )

        # Prepare data for display
        history_df = pd.DataFrame(st.session_state.history)

        # Filter
        filtered_history = history_df
        if filter_name:
            filtered_history = filtered_history[
                filtered_history["measurement_name"].str.contains(
                    filter_name, case=False, na=False
                )
            ]
        if filter_verdict != "all":
            filtered_history = filtered_history[
                filtered_history["verdict"] == filter_verdict
            ]

        # Select columns for display
        display_cols = [
            "id",
            "measurement_name",
            "verdict",
            "Cp",
            "CpkCurrent",
            "shiftValue",
            "tm",
            "lsl",
            "usl",
            "x_bar",
            "s",
            "n_samples",
            "ppm_below",
            "ppm_above",
        ]
        # Rename for clarity
        rename_map = {
            "id": "Timestamp",
            "measurement_name": "Name",
            "verdict": "Verdict",
            "CpkCurrent": "Cpk",
            "shiftValue": "Shift (Δ)",
            "tm": "Tₘ",
            "x_bar": "Mean (x̄)",
            "s": "StdDev (σ)",
            "n_samples": "n",
            "ppm_below": "PPM < LSL",
            "ppm_above": "PPM > USL",
        }

        display_df = filtered_history[display_cols].copy()
        display_df.rename(columns=rename_map, inplace=True)

        # Format for display
        format_config = {
            "Timestamp": st.column_config.TextColumn(),
            "Cp": st.column_config.NumberColumn(format="%.3f"),
            "Cpk": st.column_config.NumberColumn(format="%.3f"),
            "Shift (Δ)": st.column_config.NumberColumn(format="%.3f"),
            "Tₘ": st.column_config.NumberColumn(format="%.3f"),
            "LSL": st.column_config.NumberColumn(format="%.3f"),
            "USL": st.column_config.NumberColumn(format="%.3f"),
            "Mean (x̄)": st.column_config.NumberColumn(format="%.3f"),
            "StdDev (σ)": st.column_config.NumberColumn(format="%.4f"),
            "PPM < LSL": st.column_config.NumberColumn(format="%d"),
            "PPM > USL": st.column_config.NumberColumn(format="%d"),
        }

        st.markdown("Select rows to export:")
        selection = st.data_editor(
            display_df,
            column_config=format_config,
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",  # This allows selection
        )

        selected_indices = selection.get("selection", {}).get("rows", [])

        if selected_indices:
            selected_ids = display_df.iloc[selected_indices]["Timestamp"].tolist()
            # Find the original full data from session_state
            selected_history_data = [
                entry
                for entry in st.session_state.history
                if datetime.datetime.fromisoformat(entry["id"]).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                in [ts.strftime("%Y-%m-%d %H:%M:%S") for ts in selected_ids]
            ]

            try:
                history_buffer = exporter.export_selected_history(selected_history_data)
                with hist_cols[2]:
                    st.download_button(
                        label=f"Export Selected ({len(selected_indices)})",
                        data=history_buffer,
                        file_name=f"Capability_History_Selection_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except Exception as e:
                st.error(f"Could not generate history export: {e}")
        else:
            with hist_cols[2]:
                st.button(
                    "Export Selected (0)", use_container_width=True, disabled=True
                )


# --- Tab 5: Reference ---
with tab_ref:
    st.header("Reference Guide & Chatbot")

    ref_cols = st.columns([2, 1])

    with ref_cols[0]:
        st.subheader("Application Context & Usage Guide")
        st.markdown(
            """
            #### Technical Application: Process Centering and Root Cause Analysis
            This tool is primarily utilized in **Six Sigma and Statistical Process Control (SPC)** environments for **Process Centering and Tolerance Verification**. It enables engineers to:
            -   **Quantify Process Drift:** Calculate the exact **Required Shift (Δ)** needed to move the measured process mean (x̄) back to the engineering target (Tₘ).
            -   **Predict Initial State:** Infer the required **initial dimension/setting** to achieve the target (Tₘ).
            -   **Verify Tolerance Adequacy:** Determine the minimum **Required Tolerance** (USL - LSL) necessary for the existing process variation (σ) to meet a desired capability index (Cₚₖ, Pₚₖ).

            #### Step-by-Step Usage
            1.  **Input Data (Tabs):** Go to the **Data Worksheet** tab to paste raw data. Then, go to the **Analysis** tab to set **Specifications** (LSL/USL) and select "Use Data Worksheet".
            2.  **Run Analysis & Plot:** Click **ANALYZE & PLOT**. The tool calculates all metrics and renders plots.
            3.  **Interpret Results (Section II & III):** Review the raw numbers in Section II and read the plain-language summary in Section III.
            """
        )
        st.divider()
        st.subheader("Concepts Overview: Six Sigma & Capability Indices")
        st.markdown(
            """
            #### Six Sigma and Process Management
            **Six Sigma (6σ)** is a methodology focused on reducing process variation.
            -   3σ Level: 2,700 PPM (Defects)
            -   4σ Level: 63 PPM
            -   5σ Level: 0.57 PPM
            -   6σ Level: 0.002 PPM

            #### Capability Indices: Time Horizon
            -   **Cₚ/Cₚₖ/Cₘₖ (Short-Term Capability):** Use within-subgroup variation, reflecting immediate, inherent capability.
            -   **Pₚ/Pₚₖ (Long-Term Performance):** Use overall process variation, including all sources of variation over time.

            #### Choosing a Distribution: Normal vs. Lognormal
            -   **Normal (Gaussian):** Symmetric, bell-shaped. Use when data is symmetric (e.g., positional error).
            -   **Lognormal:** Skewed to the right, bounded by zero. Use when data cannot be negative (e.g., time to failure, concentration levels).
            """
        )
        st.divider()
        st.subheader("Additional Metric Definitions")
        st.markdown(
            """
            -   **x̄ ± 4σ Spread:** Contains ~99.9937% of process output, a conservative view of process spread.
            -   **P(x < LSL) & P(x > USL):** Probability a single part will fall outside specification limits.
            -   **P(x < Tₘ):** Probability a measurement will fall below the Target Mean. Should be 50% for a centered process.
            -   **PPM (Parts Per Million):** A measure of the defect rate (e.g., 0.001% = 10 PPM).
            """
        )
        st.divider()
        st.subheader("Hypothesis Testing & Confidence Level")
        st.markdown(
            """
            This tool performs a **Z-test** to determine if the true process mean (μ) is statistically different from the **Target Mean (Tₘ)**.
            -   **Null Hypothesis (H₀): μ = Tₘ**. (Assumes the process is on target).
            -   **p-value:** The probability of observing your sample results if H₀ were true. A small p-value (e.g., < 0.05) provides evidence against H₀.
            -   **Result: Fail to Reject H₀:** High p-value. No significant evidence the mean has shifted.
            -   **Result: Reject H₀:** Low p-value. Significant evidence the mean has shifted.

            #### When to Use Each Hypothesis Type
            -   **Two-Sided (Default):** Deviations in **both directions** are of concern.
            -   **Upper-Sided:** Critical to assess if the mean is significantly **above** the target.
            -   **Lower-Sided:** Critical to assess if the mean is significantly **below** the target.
            """
        )

    with ref_cols[1]:
        st.subheader("Guide Chatbot")
        st.info("Ask me a question about the reference guide!")

        # Chat message history
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = []

        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Chat input
        if prompt := st.chat_input("Ask about 'Cp', 'Lognormal', 'Hypothesis' ..."):
            # Add user message to chat history
            st.session_state.chat_messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            # Get bot response
            response = bot.get_response(prompt)

            # Add bot response to chat history
            with st.chat_message("assistant"):
                st.markdown(response)
            st.session_state.chat_messages.append(
                {"role": "assistant", "content": response}
            )

# --- Floating Sigma Assistant (Clippy-style) ---
# This renders as a fixed position widget in the bottom-right corner of the page
# Using st.markdown to inject directly into Streamlit's DOM for TRUE fixed positioning
mascot_html = SigmaAssistant.render_fixed(
    state=st.session_state.get("mascot_state", "idle"),
    message=st.session_state.get("mascot_message", None),
)
st.markdown(mascot_html, unsafe_allow_html=True)
