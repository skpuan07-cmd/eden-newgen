"""Extended AI tables — ai_forecast, ai_imagery_job, ai_translation_job, ai_moderation_queue."""
from datetime import datetime, date
from typing import Optional

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, String, Text, Float
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin


class AIForecast(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Predictive analytics output (Use Case 3)."""
    __tablename__ = "ai_forecast"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    target_table: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    target_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    metric: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    forecast_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    forecast_value: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    confidence_lower: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    confidence_upper: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    horizon_days: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    model_version: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    generated_on: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)


class AIImageryJob(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Satellite/drone imagery analysis (Use Case 4)."""
    __tablename__ = "ai_imagery_job"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("auth_user.id", ondelete="SET NULL"), nullable=True, index=True)
    image_url: Mapped[Optional[str]] = mapped_column(String(512), nullable=True)
    uploaded_file: Mapped[Optional[str]] = mapped_column(String(512), nullable=True)
    imagery_type: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    analysis_type: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    result_json: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    result_summary: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    location_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="pending")
    model: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)


class AITranslationJob(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Batch translation jobs (Use Case 5)."""
    __tablename__ = "ai_translation_job"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    source_lang: Mapped[str] = mapped_column(String(8), default="en")
    target_lang: Mapped[str] = mapped_column(String(8), nullable=True)
    string_key: Mapped[Optional[str]] = mapped_column(String(256), nullable=True)
    source_text: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    translated_text: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="pending")
    reviewer_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("auth_user.id", ondelete="SET NULL"), nullable=True)
    reviewed_on: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    reviewer_comment: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    model: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)


class AIModerationQueue(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Content moderation review queue (Use Case 10)."""
    __tablename__ = "ai_moderation_queue"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    source_table: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    source_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    content_preview: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    categories: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    scores: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    action: Mapped[str] = mapped_column(String(16), default="pending")
    reviewer_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("auth_user.id", ondelete="SET NULL"), nullable=True)
    reviewed_on: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    auto_action: Mapped[Optional[str]] = mapped_column(String(16), nullable=True)
    model: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
