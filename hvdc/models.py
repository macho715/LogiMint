"""HVDC 데이터 모델 정의. HVDC data model definitions."""

from __future__ import annotations

from typing import Any, Dict

from pydantic import BaseModel, ConfigDict


class LogiBaseModel(BaseModel):
    """로지스틱스 기본 모델. Logistics base model."""

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid", populate_by_name=True)

    def model_dump_safe(self) -> Dict[str, Any]:
        """안전 덤프 제공. Provide safe model dump."""
        return self.model_dump(mode="json")
