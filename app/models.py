from datetime import datetime
from typing import Optional

from sqlmodel import SQLModel, Field


class Dataset(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    name: str
    filepath: str
    uploaded_at: datetime = Field(default_factory=datetime.utcnow)
    is_active: bool = False


class Issue(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    stream: str
    field: str
    match_rate: Optional[float] = None
    not_matched: Optional[int] = None
    owner: str = ""
    status: str = "Open"
    note: str = ""
    created_at: datetime = Field(default_factory=datetime.utcnow)


class Connection(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    source: str
    target: str
    key: str = ""
    report: str = ""


class Impact(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    stream: str
    field: str
    report: str = ""
    kpi: str = ""
    owner: str = ""
    impact: str = "Low"


class AutomationRule(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    recipient: str = ""
    field: str
    condition: str
    value: str = ""
    subject: str = "Medi Intelligence Alert"
    body: str = ""
    active: bool = True
    last_run: Optional[datetime] = None
    last_hits: int = 0
