"""
Base interfaces and data structures for Commentary Providers.
Commentary must link to verified financial facts and remain provider-agnostic.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from services.finance.canonical_schema import ReportContext


@dataclass
class FactReference:
    fact_id: str
    label: str
    value: Any
    category: str = "financial"


@dataclass
class CommentarySection:
    section_id: str
    title: str
    lead_summary: str
    bullet_points: List[str] = field(default_factory=list)
    fact_references: List[FactReference] = field(default_factory=list)
    narrative_style: str = "standard"  # 'concise', 'standard', 'detailed'
    character_count: int = 0

    def __post_init__(self):
        full_text = self.lead_summary + " " + " ".join(self.bullet_points)
        self.character_count = len(full_text)


@dataclass
class PresentationCommentary:
    executive_summary: Optional[CommentarySection] = None
    performance_overview: Optional[CommentarySection] = None
    pnl_commentary: Optional[CommentarySection] = None
    balance_sheet_commentary: Optional[CommentarySection] = None
    cash_flow_commentary: Optional[CommentarySection] = None
    kpi_commentary: Optional[CommentarySection] = None
    risks_commentary: Optional[CommentarySection] = None
    actions_commentary: Optional[CommentarySection] = None
    appendix_commentary: Optional[CommentarySection] = None
    custom_sections: Dict[str, CommentarySection] = field(default_factory=dict)

    def get_section(self, section_id: str) -> Optional[CommentarySection]:
        mapping = {
            "executive_summary": self.executive_summary,
            "performance_overview": self.performance_overview,
            "pnl": self.pnl_commentary,
            "balance_sheet": self.balance_sheet_commentary,
            "cash_flow": self.cash_flow_commentary,
            "kpis": self.kpi_commentary,
            "risks": self.risks_commentary,
            "actions": self.actions_commentary,
            "appendix": self.appendix_commentary,
        }
        if section_id in mapping:
            return mapping[section_id]
        return self.custom_sections.get(section_id)


class CommentaryProvider(ABC):
    """
    Abstract interface for presentation narrative generation.
    """

    @abstractmethod
    def generate_commentary(self, context: ReportContext, style: str = "standard") -> PresentationCommentary:
        """
        Generates structured commentary sections from validated financial facts.
        """
        pass
