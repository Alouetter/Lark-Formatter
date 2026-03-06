"""规则基类定义"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from docx import Document
from src.engine.change_tracker import ChangeTracker
from src.scene.schema import SceneConfig


@dataclass
class ValidationIssue:
    level: str  # "error" | "warning" | "info"
    rule_name: str
    message: str
    location: str = ""


class BaseRule(ABC):
    name: str = ""
    description: str = ""

    @abstractmethod
    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        ...

    def validate(self, doc: Document,
                 config: SceneConfig,
                 context: dict = None) -> list[ValidationIssue]:
        return []
