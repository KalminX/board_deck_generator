"""
Native PowerPoint Chart Generator.
Creates editable python-pptx charts (Clustered Columns, Bar, and Line charts).
Applies brand colors, data labels, clean legend positioning and axis formatting.
"""
from typing import List, Dict, Any, Optional
from decimal import Decimal
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from services.presentations.tokens import BrandConfig, DEFAULT_BRAND


def create_actual_vs_budget_chart(
    slide,
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
    categories: List[str],
    actual_values: List[float],
    budget_values: List[float],
    title: str = "Actual vs Budget Comparison ($k)",
    brand: BrandConfig = DEFAULT_BRAND
):
    """
    Renders a native clustered column chart comparing Actual vs Budget.
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Actual", actual_values)
    chart_data.add_series("Budget", budget_values)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height, chart_data
    )
    chart = graphic_frame.chart

    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.chart_title.text_frame.paragraphs[0].font.name = brand.font_title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(10)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = brand.primary_color

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.name = brand.font_body
    chart.legend.font.size = Pt(8.5)

    # Series styling
    if len(chart.series) >= 2:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = brand.accent_color
        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.rgb = brand.secondary_color

    # Value Axis
    val_axis = chart.value_axis
    val_axis.has_major_gridlines = True
    val_axis.major_gridlines.format.line.color.rgb = brand.card_border
    val_axis.tick_labels.font.name = brand.font_body
    val_axis.tick_labels.font.size = Pt(8)

    # Category Axis
    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.name = brand.font_body
    cat_axis.tick_labels.font.size = Pt(8)

    return graphic_frame


def create_monthly_trend_chart(
    slide,
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
    monthly_data: List[Dict[str, Any]],
    title: str = "Historical Revenue & Profit Trend ($k)",
    brand: BrandConfig = DEFAULT_BRAND
):
    """
    Renders historical monthly revenue vs profit trend.
    """
    if not monthly_data:
        # Fallback empty chart container
        return None

    categories = [str(d.get("period", "")) for d in monthly_data]
    rev_vals = [float(d.get("revenue", 0) or 0) / 1000.0 for d in monthly_data]
    profit_vals = [float(d.get("profit", 0) or 0) / 1000.0 for d in monthly_data]

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Revenue ($k)", rev_vals)
    chart_data.add_series("Operating Profit ($k)", profit_vals)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height, chart_data
    )
    chart = graphic_frame.chart

    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.chart_title.text_frame.paragraphs[0].font.name = brand.font_title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(10)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = brand.primary_color

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.name = brand.font_body
    chart.legend.font.size = Pt(8.5)

    if len(chart.series) >= 2:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = brand.accent_color
        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.rgb = brand.success_color

    val_axis = chart.value_axis
    val_axis.has_major_gridlines = True
    val_axis.major_gridlines.format.line.color.rgb = brand.card_border
    val_axis.tick_labels.font.name = brand.font_body
    val_axis.tick_labels.font.size = Pt(8)

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.name = brand.font_body
    cat_axis.tick_labels.font.size = Pt(8)

    return graphic_frame


def create_kpi_trend_chart(
    slide,
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
    kpi_name: str,
    categories: List[str],
    values: List[float],
    targets: Optional[List[float]] = None,
    unit: str = "$",
    brand: BrandConfig = DEFAULT_BRAND
):
    """
    Renders a line chart for KPI progress over time.
    """
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Actual", values)
    if targets:
        chart_data.add_series("Target", targets)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        left, top, width, height, chart_data
    )
    chart = graphic_frame.chart

    chart.has_title = True
    chart.chart_title.text_frame.text = f"{kpi_name} Trend ({unit})"
    chart.chart_title.text_frame.paragraphs[0].font.name = brand.font_title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(9.5)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True
    chart.chart_title.text_frame.paragraphs[0].font.color.rgb = brand.primary_color

    chart.has_legend = True if targets else False
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.name = brand.font_body
        chart.legend.font.size = Pt(8)

    chart.series[0].format.line.color.rgb = brand.accent_color
    chart.series[0].format.line.width = Pt(2.5)

    if targets and len(chart.series) > 1:
        chart.series[1].format.line.color.rgb = brand.danger_color
        chart.series[1].format.line.width = Pt(1.5)

    val_axis = chart.value_axis
    val_axis.has_major_gridlines = True
    val_axis.major_gridlines.format.line.color.rgb = brand.card_border
    val_axis.tick_labels.font.name = brand.font_body
    val_axis.tick_labels.font.size = Pt(8)

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.name = brand.font_body
    cat_axis.tick_labels.font.size = Pt(8)

    return graphic_frame
