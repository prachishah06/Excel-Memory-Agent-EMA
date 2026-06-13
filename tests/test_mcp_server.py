"""Strict TDD tests for the M7 MCP server contract.

These tests define the expected behavior for the MCP layer before
`ema.mcp_server` is implemented. The MCP server must remain a thin wrapper over
`EmaService` with serializable tool outputs and structured error translation.
"""

from __future__ import annotations

import asyncio
import importlib
import json
from datetime import UTC, datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from ema import config
from ema.errors import SchemaMismatchError, ValidationError, WorkbookNotFoundError
from ema.models import (
    AppendRequest,
    AppendResult,
    ColumnDef,
    ColumnType,
    SchemaProposal,
    UndoResult,
    WorkbookEntry,
    WorkbookSchema,
)


def _load_mcp_module():
    """Import the future M7 MCP module under test."""

    return importlib.import_module("ema.mcp_server")


def _make_server(service: object):
    """Construct the MCP server wrapper with an injected service instance."""

    module = _load_mcp_module()
    return module.McpServer(service=service)


def _make_fastmcp_server(service: object):
    """Construct the real FastMCP server with an injected service instance."""

    module = _load_mcp_module()
    return module.create_mcp_server(service=service)


def _registered_tools(mcp_server: FastMCP):
    """Return the registered FastMCP tools using the public async API."""

    return asyncio.run(mcp_server.list_tools())


def _example_schema(*, table_name: str | None = None) -> WorkbookSchema:
    """Build a stable schema object for MCP serialization tests."""

    return WorkbookSchema(
        schema_version=config.SCHEMA_VERSION,
        sheet="Expenses" if table_name else "FoodLog",
        header_row=1,
        table_name=table_name,
        columns=[
            ColumnDef(name="Date", index=0, type=ColumnType.DATE, required=True),
            ColumnDef(
                name="Store" if table_name else "Meal",
                index=1,
                type=ColumnType.TEXT,
                required=True,
            ),
            ColumnDef(
                name="Amount" if table_name else "Calories",
                index=2,
                type=ColumnType.NUMBER,
                required=True,
            ),
        ],
        data_start_row=2,
    )


def _example_entry() -> WorkbookEntry:
    """Build a stable registered workbook entry for MCP list serialization tests."""

    now = datetime(2026, 6, 12, 12, 0, 0, tzinfo=UTC)
    return WorkbookEntry(
        id="expenses",
        name="Business Expenses",
        path=Path("C:/workbooks/expenses.xlsx"),
        primary_sheet="Expenses",
        schema=_example_schema(table_name="ExpensesTable"),
        created_at=now,
        updated_at=now,
        last_undo_token="backup-token",
    )


class FakeService:
    """Minimal fake service that records calls for MCP delegation tests."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[object, ...], dict[str, object]]] = []
        self.register_result: object = SchemaProposal(
            sheet="FoodLog",
            header_row=1,
            table_name=None,
            columns=_example_schema().columns,
            warnings=[],
        )
        self.list_result: list[WorkbookEntry] = [_example_entry()]
        self.schema_result: WorkbookSchema = _example_schema(table_name="ExpensesTable")
        self.append_result: AppendResult = AppendResult(
            ok=True,
            workbook_id="expenses",
            sheet="Expenses",
            written_row=5,
            row_preview={"Date": "2026-06-12", "Store": "Lidl", "Amount": 22.75},
            backup_path="C:/backups/expenses.xlsx",
            undo_token="backup-token",
            dry_run=False,
            message="Appended row 5.",
        )
        self.undo_result: UndoResult = UndoResult(
            ok=True,
            workbook_id="expenses",
            restored_from="C:/backups/expenses.xlsx",
            message="Restored workbook from backup.",
        )

    def register_workbook(
        self,
        path: str,
        name: str | None = None,
        sheet: str | None = None,
        header_row: int | None = None,
        confirm: bool = True,
    ) -> object:
        self.calls.append(
            (
                "register_workbook",
                (path,),
                {
                    "name": name,
                    "sheet": sheet,
                    "header_row": header_row,
                    "confirm": confirm,
                },
            )
        )
        return self.register_result

    def list_workbooks(self) -> list[WorkbookEntry]:
        self.calls.append(("list_workbooks", (), {}))
        return self.list_result

    def get_workbook_schema(self, workbook_id: str) -> WorkbookSchema:
        self.calls.append(("get_workbook_schema", (workbook_id,), {}))
        return self.schema_result

    def append_row(self, req: AppendRequest) -> AppendResult:
        self.calls.append(("append_row", (req,), {}))
        return self.append_result

    def undo_last_append(self, workbook_id: str) -> UndoResult:
        self.calls.append(("undo_last_append", (workbook_id,), {}))
        return self.undo_result


def test_register_workbook_delegates_to_service_register_workbook() -> None:
    """Purpose: prove the MCP tool is only a thin wrapper over the service.

    Expected behavior: the tool forwards workbook registration arguments directly
    to `EmaService.register_workbook(...)` and returns a JSON-serializable
    structure built from the service result.

    Why it exists: M7 must not move registration logic into the MCP layer.
    """

    service = FakeService()
    server = _make_server(service)

    result = server.register_workbook(
        path="C:/workbooks/foodlog.xlsx",
        name="Food Log",
        sheet="FoodLog",
        header_row=1,
        confirm=False,
    )

    assert service.calls == [
        (
            "register_workbook",
            ("C:/workbooks/foodlog.xlsx",),
            {"name": "Food Log", "sheet": "FoodLog", "header_row": 1, "confirm": False},
        )
    ]
    assert json.loads(json.dumps(result)) == result


def test_list_workbooks_delegates_to_service_list_workbooks() -> None:
    """Purpose: ensure the list tool simply exposes the service registry view.

    Expected behavior: the tool calls `EmaService.list_workbooks()` exactly once
    and returns a JSON-serializable dictionary structure.

    Why it exists: the MCP layer must not query registry state independently.
    """

    service = FakeService()
    server = _make_server(service)

    result = server.list_workbooks()

    assert service.calls == [("list_workbooks", (), {})]
    assert json.loads(json.dumps(result)) == result


def test_get_workbook_schema_delegates_to_service_get_workbook_schema() -> None:
    """Purpose: prove schema retrieval flows through the service boundary.

    Expected behavior: the tool forwards the workbook identifier to
    `EmaService.get_workbook_schema(...)` and serializes the returned schema.

    Why it exists: persisted schema access belongs to M6 service logic, not M7.
    """

    service = FakeService()
    server = _make_server(service)

    result = server.get_workbook_schema("expenses")

    assert service.calls == [("get_workbook_schema", ("expenses",), {})]
    assert json.loads(json.dumps(result)) == result


def test_append_row_delegates_to_service_append_row() -> None:
    """Purpose: lock the append tool to the single public service API.

    Expected behavior: the tool constructs one `AppendRequest`, passes it to
    `EmaService.append_row(...)`, and returns a serializable dict.

    Why it exists: callers must never reach writer helpers or re-implement
    append orchestration in the MCP layer.
    """

    service = FakeService()
    server = _make_server(service)

    result = server.append_row(
        workbook_id="expenses",
        values={"Date": "2026-06-12", "Store": "Lidl", "Amount": 22.75},
        sheet="Expenses",
        dry_run=False,
    )

    assert len(service.calls) == 1
    method, args, kwargs = service.calls[0]
    assert method == "append_row"
    assert kwargs == {}
    assert len(args) == 1
    assert isinstance(args[0], AppendRequest)
    assert args[0].workbook_id == "expenses"
    assert args[0].sheet == "Expenses"
    assert args[0].values == {"Date": "2026-06-12", "Store": "Lidl", "Amount": 22.75}
    assert args[0].dry_run is False
    assert json.loads(json.dumps(result)) == result


def test_undo_last_append_delegates_to_service_undo_last_append() -> None:
    """Purpose: prove undo goes through the service token-tracking path.

    Expected behavior: the tool forwards only the workbook identifier to
    `EmaService.undo_last_append(...)` and returns a serializable dict.

    Why it exists: the MCP layer must not resolve backup tokens or perform undo
    bookkeeping itself.
    """

    service = FakeService()
    server = _make_server(service)

    result = server.undo_last_append("expenses")

    assert service.calls == [("undo_last_append", ("expenses",), {})]
    assert json.loads(json.dumps(result)) == result


def test_workbook_not_found_error_is_translated_to_structured_error_dict() -> None:
    """Purpose: verify not-found service errors become MCP-safe structured output.

    Expected behavior: `WorkbookNotFoundError` is not re-raised; instead the tool
    returns `{"ok": False, "error": ..., "message": ...}`.

    Why it exists: MCP tools must surface business errors as protocol-safe data.
    """

    service = FakeService()

    def raise_missing(_workbook_id: str) -> WorkbookSchema:
        raise WorkbookNotFoundError("Workbook is not registered: missing")

    service.get_workbook_schema = raise_missing  # type: ignore[method-assign]
    server = _make_server(service)

    result = server.get_workbook_schema("missing")

    assert result == {
        "ok": False,
        "error": "WorkbookNotFoundError",
        "message": "Workbook is not registered: missing",
    }


def test_validation_error_is_translated_to_structured_error_dict() -> None:
    """Purpose: verify validation failures remain thinly translated at the MCP layer.

    Expected behavior: `ValidationError` from the service is returned as a
    structured error dict rather than being wrapped in new business logic.

    Why it exists: M7 should translate, not reinterpret, service validation.
    """

    service = FakeService()

    def raise_validation(_req: AppendRequest) -> AppendResult:
        raise ValidationError("Unknown column(s) for sheet 'Expenses': ['Currency']")

    service.append_row = raise_validation  # type: ignore[method-assign]
    server = _make_server(service)

    result = server.append_row(
        workbook_id="expenses",
        values={"Date": "2026-06-12", "Store": "Lidl", "Amount": 22.75, "Currency": "EUR"},
    )

    assert result == {
        "ok": False,
        "error": "ValidationError",
        "message": "Unknown column(s) for sheet 'Expenses': ['Currency']",
    }


def test_schema_mismatch_error_is_translated_to_structured_error_dict() -> None:
    """Purpose: verify drift-detection failures survive the MCP translation layer.

    Expected behavior: `SchemaMismatchError` becomes the standard structured
    error dict and is not swallowed or transformed into success output.

    Why it exists: clients must be able to detect schema drift reliably.
    """

    service = FakeService()

    def raise_drift(_req: AppendRequest) -> AppendResult:
        raise SchemaMismatchError("Header drift detected for sheet 'FoodLog'.")

    service.append_row = raise_drift  # type: ignore[method-assign]
    server = _make_server(service)

    result = server.append_row(
        workbook_id="foodlog",
        values={"Date": "2026-06-12", "Meal": "Snack", "Calories": 95},
    )

    assert result == {
        "ok": False,
        "error": "SchemaMismatchError",
        "message": "Header drift detected for sheet 'FoodLog'.",
    }


def test_returned_values_are_json_serializable_structures() -> None:
    """Purpose: ensure every MCP tool result is transport-safe plain data.

    Expected behavior: successful tool outputs can be round-tripped through
    `json.dumps(...)` and `json.loads(...)` without custom encoders.

    Why it exists: MCP tool results must cross process boundaries cleanly.
    """

    service = FakeService()
    server = _make_server(service)

    results = [
        server.register_workbook(path="C:/workbooks/foodlog.xlsx", confirm=False),
        server.list_workbooks(),
        server.get_workbook_schema("expenses"),
        server.append_row(
            workbook_id="expenses",
            values={"Date": "2026-06-12", "Store": "Lidl", "Amount": 22.75},
        ),
        server.undo_last_append("expenses"),
    ]

    for result in results:
        assert json.loads(json.dumps(result)) == result


def test_mcp_server_can_be_constructed_with_injected_service() -> None:
    """Purpose: lock in dependency injection for MCP/server tests and composition.

    Expected behavior: the MCP server is constructible with an explicit
    `EmaService` instance and exposes the five always-on tools.

    Why it exists: M7 must stay testable and must not hardcode its own service.
    """

    service = FakeService()
    server = _make_server(service)

    assert hasattr(server, "register_workbook")
    assert hasattr(server, "list_workbooks")
    assert hasattr(server, "get_workbook_schema")
    assert hasattr(server, "append_row")
    assert hasattr(server, "undo_last_append")


def test_create_mcp_server_returns_fastmcp_instance() -> None:
    """Purpose: lock the public MCP entry point to a real FastMCP server.

    Expected behavior: `create_mcp_server(...)` returns a `FastMCP` instance
    rather than the lower-level adapter object.

    Why it exists: the architecture requires a real FastMCP layer above
    `McpServer` for host integration.
    """

    service = FakeService()

    mcp_server = _make_fastmcp_server(service)

    assert isinstance(mcp_server, FastMCP)


def test_create_mcp_server_registers_exactly_five_tools() -> None:
    """Purpose: prevent accidental exposure of extra or missing structured tools.

    Expected behavior: the FastMCP server registers exactly five always-on tools.

    Why it exists: M7 should expose only the structured path defined by the
    architecture at this stage.
    """

    service = FakeService()

    mcp_server = _make_fastmcp_server(service)
    tools = _registered_tools(mcp_server)

    assert len(tools) == 5


def test_create_mcp_server_registers_expected_tool_names() -> None:
    """Purpose: freeze the public MCP tool names used by hosts.

    Expected behavior: the FastMCP server registers exactly these names:
    `register_workbook`, `list_workbooks`, `get_workbook_schema`,
    `append_row`, and `undo_last_append`.

    Why it exists: host-facing tool names are part of the external contract.
    """

    service = FakeService()

    mcp_server = _make_fastmcp_server(service)
    tools = _registered_tools(mcp_server)

    assert [tool.name for tool in tools] == [
        "register_workbook",
        "list_workbooks",
        "get_workbook_schema",
        "append_row",
        "undo_last_append",
    ]


def test_registered_fastmcp_tools_delegate_into_mcp_server_methods() -> None:
    """Purpose: prove FastMCP tools are bound to the adapter layer, not service methods.

    Expected behavior: each registered tool wraps a bound `McpServer` method
    whose function name matches the public tool name.

    Why it exists: the FastMCP layer must route through `McpServer`, keeping the
    server composition `FastMCP -> McpServer -> EmaService`.
    """

    module = _load_mcp_module()
    service = FakeService()

    mcp_server = _make_fastmcp_server(service)
    tools = _registered_tools(mcp_server)

    for tool in tools:
        assert tool.fn.__self__.__class__ is module.McpServer
        assert tool.fn.__name__ == tool.name


def test_create_mcp_server_preserves_injected_service_inside_adapter() -> None:
    """Purpose: ensure dependency injection survives FastMCP registration.

    Expected behavior: every registered tool remains bound to the same
    `McpServer` instance, and that adapter holds the exact injected service.

    Why it exists: M7 must stay testable and must not silently replace the
    caller-supplied service dependency.
    """

    service = FakeService()

    mcp_server = _make_fastmcp_server(service)
    tools = _registered_tools(mcp_server)
    adapter_instances = {id(tool.fn.__self__): tool.fn.__self__ for tool in tools}

    assert len(adapter_instances) == 1
    adapter = next(iter(adapter_instances.values()))
    assert adapter._service is service


def test_fastmcp_layer_does_not_bypass_mcp_server() -> None:
    """Purpose: prevent direct FastMCP-to-service coupling.

    Expected behavior: registered tool callables are bound to `McpServer`
    instances and not directly to the injected service object.

    Why it exists: M7 must remain a thin two-step adapter, not collapse the MCP
    layer directly onto `EmaService`.
    """

    service = FakeService()

    mcp_server = _make_fastmcp_server(service)
    tools = _registered_tools(mcp_server)

    for tool in tools:
        assert tool.fn.__self__ is not service
