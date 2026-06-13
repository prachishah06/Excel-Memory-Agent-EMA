"""Thin MCP adapter over the structured EMA service layer."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from ema.errors import EmaError
from ema.models import AppendRequest
from ema.service import EmaService


class EmaFastMCP(FastMCP):
    """Small FastMCP subclass that exposes bound tool objects for inspection."""

    async def list_tools(self):
        """Return the registered bound tool objects in registration order."""

        return list(self._tool_manager._tools.values())


class McpServer:
    """Expose the structured EMA service API as MCP-friendly tool methods."""

    def __init__(self, service: EmaService | None = None) -> None:
        """Create an MCP adapter with an injectable service dependency."""

        self._service = service or EmaService()

    def register_workbook(
        self,
        path: str,
        name: str | None = None,
        sheet: str | None = None,
        header_row: int | None = None,
        confirm: bool = False,
    ) -> dict:
        """Delegate workbook registration to the structured service layer."""

        try:
            result = self._service.register_workbook(path, name, sheet, header_row, confirm)
            return result.model_dump(mode="json")
        except EmaError as exc:
            return _error_dict(exc)

    def list_workbooks(self) -> dict:
        """Return registered workbooks as a JSON-serializable structure."""

        try:
            workbooks = self._service.list_workbooks()
            return {"workbooks": [entry.model_dump(mode="json") for entry in workbooks]}
        except EmaError as exc:
            return _error_dict(exc)

    def get_workbook_schema(self, workbook_id: str) -> dict:
        """Return the persisted schema for a registered workbook."""

        try:
            result = self._service.get_workbook_schema(workbook_id)
            return result.model_dump(mode="json")
        except EmaError as exc:
            return _error_dict(exc)

    def append_row(
        self,
        workbook_id: str,
        values: dict,
        sheet: str | None = None,
        dry_run: bool = False,
    ) -> dict:
        """Delegate one structured append request to the service layer."""

        try:
            request = AppendRequest(
                workbook_id=workbook_id,
                sheet=sheet,
                values=values,
                dry_run=dry_run,
            )
            result = self._service.append_row(request)
            return result.model_dump(mode="json")
        except EmaError as exc:
            return _error_dict(exc)

    def undo_last_append(self, workbook_id: str) -> dict:
        """Delegate undo of the most recent append to the service layer."""

        try:
            result = self._service.undo_last_append(workbook_id)
            return result.model_dump(mode="json")
        except EmaError as exc:
            return _error_dict(exc)


def create_mcp_server(service: EmaService | None = None) -> FastMCP:
    """Construct a real FastMCP server bound to one injected MCP adapter."""

    adapter = McpServer(service=service)
    mcp = EmaFastMCP("excel-memory-agent")

    mcp.tool(name="register_workbook")(adapter.register_workbook)
    mcp.tool(name="list_workbooks")(adapter.list_workbooks)
    mcp.tool(name="get_workbook_schema")(adapter.get_workbook_schema)
    mcp.tool(name="append_row")(adapter.append_row)
    mcp.tool(name="undo_last_append")(adapter.undo_last_append)

    return mcp


def _error_dict(exc: EmaError) -> dict[str, object]:
    """Translate an EMA exception into an MCP-safe structured error payload."""

    return {"ok": False, "error": type(exc).__name__, "message": str(exc)}


__all__ = ["EmaFastMCP", "McpServer", "create_mcp_server"]
