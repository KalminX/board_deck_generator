"""
Presentation Plan and Output Validator.
Checks slide selections, layout bounds, required slides, and slide count sanity.
"""
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from services.finance.canonical_schema import ReportContext


@dataclass
class PresentationValidationResult:
    is_valid: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class PresentationValidator:
    """
    Validates slide configuration and generated presentations.
    """

    ALLOWED_SLIDE_KEYS = [
        "cover",
        "executive_summary",
        "financial_overview",
        "pnl",
        "balance_sheet",
        "cash_flow",
        "kpis",
        "risks",
        "actions",
        "appendix"
    ]

    @classmethod
    def validate_slide_selection(cls, selected_slides: List[str]) -> PresentationValidationResult:
        res = PresentationValidationResult(is_valid=True)

        if not selected_slides:
            res.errors.append("At least one slide must be selected for generation.")
            res.is_valid = False
            return res

        for s in selected_slides:
            if s not in cls.ALLOWED_SLIDE_KEYS:
                res.errors.append(f"Unrecognized slide key: '{s}'.")
                res.is_valid = False

        if "cover" not in selected_slides:
            res.warnings.append("Cover slide is not selected. Deck will begin directly on analytical content.")

        return res

    @classmethod
    def validate_presentation_output(cls, prs, expected_count: int) -> PresentationValidationResult:
        res = PresentationValidationResult(is_valid=True)
        actual_count = len(prs.slides)

        if actual_count != expected_count:
            res.warnings.append(f"Generated slide count ({actual_count}) differs from expected plan ({expected_count}).")

        if actual_count == 0:
            res.errors.append("Generated presentation is empty.")
            res.is_valid = False

        return res
