"""Extended Eden modules — inv, cr, dvr, cap, cms, asset, req, project, vehicle, vol, budget."""
from datetime import date, datetime
from typing import Optional

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, String, Text, Float
from sqlalchemy.orm import Mapped, mapped_column

from app.models.base import Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin


class InvWarehouse(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Inventory warehouse."""
    __tablename__ = "inv_warehouse"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    organisation_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation.id"), nullable=True, index=True)
    site_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_site.id"), nullable=True)
    location_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("gis_location.id"), nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class InvInvItem(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Inventory item in a warehouse."""
    __tablename__ = "inv_inv_item"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    warehouse_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("inv_warehouse.id", ondelete="SET NULL"), nullable=True, index=True)
    quantity: Mapped[int] = mapped_column(Integer, default=0)
    unit: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="in_stock")
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class CrShelter(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Shelter (ศูนย์พักพิง)."""
    __tablename__ = "cr_shelter"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    site_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_site.id"), nullable=True)
    population: Mapped[int] = mapped_column(Integer, default=0)
    capacity: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="open")
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class DvrCase(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Disaster victim registry case (SENSITIVE — PDPA)."""
    __tablename__ = "dvr_case"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    person_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("pr_person.id", ondelete="SET NULL"), nullable=True, index=True)
    shelter_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("cr_shelter.id", ondelete="SET NULL"), nullable=True)
    case_status: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)
    notes_encrypted: Mapped[Optional[str]] = mapped_column(Text, nullable=True)


class CapAlert(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """CAP alert (Common Alerting Protocol)."""
    __tablename__ = "cap_alert"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    identifier: Mapped[Optional[str]] = mapped_column(String(128), nullable=True, index=True)
    sender: Mapped[Optional[str]] = mapped_column(String(128), nullable=True)
    sent: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    status: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    msg_type: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    scope: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    event: Mapped[Optional[str]] = mapped_column(String(256), nullable=True)
    headline: Mapped[Optional[str]] = mapped_column(String(512), nullable=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    severity: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)


class CmsPost(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """CMS post (news, announcements)."""
    __tablename__ = "cms_post"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    title: Mapped[str] = mapped_column(String(256), nullable=False, index=True)
    body: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    author_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("auth_user.id", ondelete="SET NULL"), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="draft")
    published_on: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class AssetAsset(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Asset (ครุภัณฑ์)."""
    __tablename__ = "asset_asset"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    asset_type: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    site_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_site.id"), nullable=True)
    assigned_to_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("hrm_human_resource.id", ondelete="SET NULL"), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="available")
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class ReqReq(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Request (คำขอ)."""
    __tablename__ = "req_req"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    person_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("pr_person.id", ondelete="SET NULL"), nullable=True)
    site_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_site.id"), nullable=True)
    req_type: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="pending")
    priority: Mapped[Optional[str]] = mapped_column(String(16), nullable=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class ProjectProject(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Project (โครงการ)."""
    __tablename__ = "project_project"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    organisation_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation.id"), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="active")
    start_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    end_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class VehicleVehicle(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Vehicle (ยานพาหนะ)."""
    __tablename__ = "vehicle_vehicle"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    vehicle_type: Mapped[Optional[str]] = mapped_column(String(64), nullable=True)
    registration: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    organisation_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation.id"), nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="available")


class VolVolunteer(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Volunteer (อาสาสมัคร)."""
    __tablename__ = "vol_volunteer"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    human_resource_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("hrm_human_resource.id", ondelete="CASCADE"), nullable=True, index=True)
    application_status: Mapped[str] = mapped_column(String(16), default="pending")
    cluster_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)


class BudgetBudget(Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin):
    """Budget (งบประมาณ)."""
    __tablename__ = "budget_budget"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    organisation_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation.id"), nullable=True)
    amount: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    currency: Mapped[str] = mapped_column(String(8), default="THB")
    status: Mapped[str] = mapped_column(String(16), default="draft")
    event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("event_event.id", ondelete="SET NULL"), nullable=True, index=True)
