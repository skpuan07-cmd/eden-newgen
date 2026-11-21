"""SQLAlchemy 2.0 models for Sahana Eden FastAPI app."""
from app.models.base import Base, UUIDMixin, MCIMixin, TimestampMixin, SoftDeleteMixin
from app.models.gis import Location
from app.models.auth import User, Group, Membership
from app.models.person import Person, PersonContact, PersonAddress
from app.models.organisation import Organisation, OrganisationType, Site
from app.models.hrm import HumanResource, JobTitle
from app.models.event import Event
from app.models.ai import AIChatLog, AIOcrJob, AIVoiceLog, AIKnowledgeChunk, AIQueryLog
from app.models.ai_extended import AIForecast, AIImageryJob, AITranslationJob, AIModerationQueue
from app.models.eden_extended import (
    InvWarehouse, InvInvItem, CrShelter, DvrCase, CapAlert, CmsPost,
    AssetAsset, ReqReq, ProjectProject, VehicleVehicle, VolVolunteer, BudgetBudget,
)

__all__ = [
    "Base", "UUIDMixin", "MCIMixin", "TimestampMixin", "SoftDeleteMixin",
    "Location", "User", "Group", "Membership",
    "Person", "PersonContact", "PersonAddress",
    "Organisation", "OrganisationType", "Site",
    "HumanResource", "JobTitle", "Event",
    "AIChatLog", "AIOcrJob", "AIVoiceLog", "AIKnowledgeChunk", "AIQueryLog",
    "AIForecast", "AIImageryJob", "AITranslationJob", "AIModerationQueue",
    "InvWarehouse", "InvInvItem", "CrShelter", "DvrCase", "CapAlert", "CmsPost",
    "AssetAsset", "ReqReq", "ProjectProject", "VehicleVehicle", "VolVolunteer", "BudgetBudget",
]
