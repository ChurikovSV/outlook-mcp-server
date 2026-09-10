from pydantic import BaseModel, Field


class TableBlock(BaseModel):
    title: str | None = None
    columns: list[str]
    rows: list[list[str]]


class UploadedAttachment(BaseModel):
    filename: str
    content_base64: str


class EmailRequest(BaseModel):
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    recipient_file: str | None = None
    recipient_file_column: str = "email"
    recipient_file_sheet: str | None = None
    subject: str
    body: str = ""
    tables: list[TableBlock] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    uploaded_attachments: list[UploadedAttachment] = Field(default_factory=list)


class BulkEmailRequest(BaseModel):
    recipients: list[str] = Field(default_factory=list)
    recipient_file: str | None = None
    recipient_file_column: str = "email"
    recipient_file_sheet: str | None = None
    subject: str
    body: str = ""
    tables: list[TableBlock] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    uploaded_attachments: list[UploadedAttachment] = Field(default_factory=list)


class DraftItem(BaseModel):
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    subject: str
    body: str = ""
    tables: list[TableBlock] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    uploaded_attachments: list[UploadedAttachment] = Field(default_factory=list)


class BatchDraftRequest(BaseModel):
    drafts: list[DraftItem]
