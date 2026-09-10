from pydantic import BaseModel, Field


class TableBlock(BaseModel):
    title: str | None = None
    columns: list[str]
    rows: list[list[str]]


class EmailRequest(BaseModel):
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    subject: str
    body: str = ""
    tables: list[TableBlock] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
