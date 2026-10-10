"""
Presentation Design Tokens and Brand Configuration.
Configurable design parameters for corporate board presentations.
Standard 16:9 widescreen layout with professional corporate color palettes.
"""
from dataclasses import dataclass
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


@dataclass
class BrandConfig:
    brand_name: str = "Corporate Slate"
    primary_color: RGBColor = RGBColor(15, 23, 42)       # Slate 900 #0F172A
    secondary_color: RGBColor = RGBColor(51, 65, 85)     # Slate 700 #334155
    accent_color: RGBColor = RGBColor(37, 99, 235)       # Royal Blue #2563EB
    accent_light: RGBColor = RGBColor(239, 246, 255)     # Blue 50 #EFF6FF
    success_color: RGBColor = RGBColor(22, 163, 74)      # Green 600 #16A34A
    danger_color: RGBColor = RGBColor(220, 38, 38)       # Red 600 #DC2626
    warning_color: RGBColor = RGBColor(217, 119, 6)      # Amber 600 #D97706
    bg_color: RGBColor = RGBColor(248, 250, 252)         # Slate 50 #F8FAFC
    card_bg: RGBColor = RGBColor(255, 255, 255)          # White #FFFFFF
    card_border: RGBColor = RGBColor(226, 232, 240)      # Slate 200 #E2E8F0
    text_dark: RGBColor = RGBColor(15, 23, 42)           # Slate 900
    text_muted: RGBColor = RGBColor(100, 116, 139)       # Slate 500 #64748B
    table_header_bg: RGBColor = RGBColor(15, 23, 42)     # Dark Slate
    table_alt_row: RGBColor = RGBColor(248, 250, 252)    # Slate 50

    font_title: str = "Calibri"
    font_body: str = "Calibri"
    font_mono: str = "Consolas"

    # Slide Dimensions (16:9 Widescreen)
    slide_width: Inches = Inches(13.333)
    slide_height: Inches = Inches(7.5)

    # Standard Margins
    margin_left: Inches = Inches(0.6)
    margin_right: Inches = Inches(0.6)
    margin_top: Inches = Inches(0.5)
    margin_bottom: Inches = Inches(0.4)

    # Reusable Typography Sizes
    size_slide_title: Pt = Pt(20)
    size_slide_subtitle: Pt = Pt(11)
    size_card_title: Pt = Pt(12)
    size_metric_value: Pt = Pt(22)
    size_metric_label: Pt = Pt(9.5)
    size_body: Pt = Pt(10)
    size_caption: Pt = Pt(8.5)
    size_table_cell: Pt = Pt(9)
    size_table_header: Pt = Pt(9.5)


DEFAULT_BRAND = BrandConfig()

# Alternate preset brand themes
MODERN_NAVY_BRAND = BrandConfig(
    brand_name="Executive Navy",
    primary_color=RGBColor(10, 25, 47),
    secondary_color=RGBColor(30, 41, 59),
    accent_color=RGBColor(2, 132, 199),
    table_header_bg=RGBColor(10, 25, 47)
)

MINIMAL_EMERALD_BRAND = BrandConfig(
    brand_name="Emerald Finance",
    primary_color=RGBColor(6, 78, 59),
    secondary_color=RGBColor(15, 118, 110),
    accent_color=RGBColor(5, 150, 105),
    table_header_bg=RGBColor(6, 78, 59)
)
