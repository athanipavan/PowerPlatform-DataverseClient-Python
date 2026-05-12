# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

import pytest
from contextlib import asynccontextmanager
from unittest.mock import AsyncMock, MagicMock, patch

from azure.core.credentials_async import AsyncTokenCredential

from PowerPlatform.Dataverse.aio.async_client import AsyncDataverseClient
from PowerPlatform.Dataverse.aio.operations.async_query import AsyncQueryOperations
from PowerPlatform.Dataverse.models.record import Record


def _make_async_client_with_od(mock_od):
    """Helper: create async client with mocked _scoped_odata."""
    cred = MagicMock(spec=AsyncTokenCredential)
    client = AsyncDataverseClient("https://example.crm.dynamics.com", cred)

    @asynccontextmanager
    async def _fake_scoped():
        yield mock_od

    client._scoped_odata = _fake_scoped
    return client


class TestAsyncQueryOperationsNamespace:
    def test_namespace_type(self, async_client):
        assert isinstance(async_client.query, AsyncQueryOperations)

    def test_no_builder_attribute(self, async_client):
        """builder() is intentionally absent — QueryBuilder.execute() is sync-only."""
        assert not hasattr(async_client.query, "builder")


class TestAsyncQuerySql:
    async def test_sql_returns_records(self, async_client, mock_od):
        """sql() calls _query_sql and wraps results in Record objects."""
        mock_od._query_sql.return_value = [
            {"name": "Contoso", "accountid": "guid-1"},
            {"name": "Fabrikam", "accountid": "guid-2"},
        ]

        result = await async_client.query.sql("SELECT TOP 2 name FROM account")

        mock_od._query_sql.assert_called_once_with("SELECT TOP 2 name FROM account")
        assert len(result) == 2
        assert all(isinstance(r, Record) for r in result)
        assert result[0]["name"] == "Contoso"
        assert result[1]["name"] == "Fabrikam"

    async def test_sql_empty_result(self, async_client, mock_od):
        """sql() returns an empty list when no rows match."""
        mock_od._query_sql.return_value = []
        result = await async_client.query.sql("SELECT name FROM account WHERE name = 'X'")
        assert result == []


class TestAsyncQuerySqlColumns:
    async def test_sql_columns_filters_virtual_and_system(self, async_client, mock_od):
        """sql_columns() calls tables.list_columns and filters out virtual/system columns."""
        mock_od._list_columns.return_value = [
            {
                "LogicalName": "name",
                "AttributeType": "String",
                "IsPrimaryId": False,
                "IsPrimaryName": True,
                "DisplayName": {},
                "AttributeOf": None,
            },
            {
                "LogicalName": "accountid",
                "AttributeType": "Uniqueidentifier",
                "IsPrimaryId": True,
                "IsPrimaryName": False,
                "DisplayName": {},
                "AttributeOf": None,
            },
            {
                "LogicalName": "versionnumber",
                "AttributeType": "BigInt",
                "IsPrimaryId": False,
                "IsPrimaryName": False,
                "DisplayName": {},
                "AttributeOf": None,
            },
        ]

        cols = await async_client.query.sql_columns("account")

        # versionnumber is a system column — excluded by default
        names = [c["name"] for c in cols]
        assert "versionnumber" not in names
        assert "accountid" in names
        assert "name" in names

    async def test_sql_columns_include_system(self, async_client, mock_od):
        """sql_columns(include_system=True) includes system columns."""
        mock_od._list_columns.return_value = [
            {
                "LogicalName": "versionnumber",
                "AttributeType": "BigInt",
                "IsPrimaryId": False,
                "IsPrimaryName": False,
                "DisplayName": {},
                "AttributeOf": None,
            }
        ]

        cols = await async_client.query.sql_columns("account", include_system=True)
        assert any(c["name"] == "versionnumber" for c in cols)

    async def test_sql_columns_excludes_attribute_of(self, async_client, mock_od):
        """sql_columns() excludes columns where AttributeOf is set."""
        mock_od._list_columns.return_value = [
            {
                "LogicalName": "parentcustomeridname",
                "AttributeType": "String",
                "IsPrimaryId": False,
                "IsPrimaryName": False,
                "DisplayName": {},
                "AttributeOf": "parentcustomerid",
            }
        ]

        cols = await async_client.query.sql_columns("contact")
        assert cols == []

    async def test_sql_columns_skips_empty_logical_name(self, async_client, mock_od):
        """sql_columns() skips columns where LogicalName is empty."""
        mock_od._list_columns.return_value = [
            {
                "LogicalName": "",
                "AttributeType": "String",
                "IsPrimaryId": False,
                "IsPrimaryName": False,
                "DisplayName": {},
                "AttributeOf": None,
            },
            {
                "LogicalName": "name",
                "AttributeType": "String",
                "IsPrimaryId": False,
                "IsPrimaryName": True,
                "DisplayName": {},
                "AttributeOf": None,
            },
        ]
        cols = await async_client.query.sql_columns("account")
        names = [c["name"] for c in cols]
        assert "" not in names
        assert "name" in names

    async def test_sql_columns_extracts_display_label(self, async_client, mock_od):
        """sql_columns() extracts label from UserLocalizedLabel when present."""
        mock_od._list_columns.return_value = [
            {
                "LogicalName": "name",
                "AttributeType": "String",
                "IsPrimaryId": False,
                "IsPrimaryName": True,
                "DisplayName": {"UserLocalizedLabel": {"Label": "Account Name", "LanguageCode": 1033}},
                "AttributeOf": None,
            },
        ]
        cols = await async_client.query.sql_columns("account")
        assert len(cols) == 1
        assert cols[0]["label"] == "Account Name"


class TestAsyncQueryOdataExpands:
    async def test_odata_expands_returns_nav_properties(self, async_client, mock_od):
        """odata_expands() returns navigation property metadata."""
        mock_od._list_table_relationships.return_value = [
            {
                "ReferencingEntity": "contact",
                "ReferencingEntityNavigationPropertyName": "parentcustomerid_account",
                "ReferencedEntity": "account",
                "ReferencingAttribute": "parentcustomerid",
                "SchemaName": "contact_customer_accounts",
            }
        ]
        mock_od._entity_set_from_schema_name.return_value = "accounts"

        result = await async_client.query.odata_expands("contact")

        assert len(result) == 1
        assert result[0]["nav_property"] == "parentcustomerid_account"
        assert result[0]["target_table"] == "account"

    async def test_odata_expands_filters_non_referencing(self, async_client, mock_od):
        """odata_expands() skips relationships where ReferencingEntity != table."""
        mock_od._list_table_relationships.return_value = [
            {
                "ReferencingEntity": "account",  # not "contact"
                "ReferencingEntityNavigationPropertyName": "ownerid_systemuser",
                "ReferencedEntity": "systemuser",
                "ReferencingAttribute": "ownerid",
                "SchemaName": "account_owner_rel",
            }
        ]
        mock_od._entity_set_from_schema_name.return_value = "systemusers"

        result = await async_client.query.odata_expands("contact")
        assert result == []

    async def test_odata_expands_skips_empty_nav_prop(self, async_client, mock_od):
        """odata_expands() skips relationships with empty nav_prop or target."""
        mock_od._list_table_relationships.return_value = [
            {
                "ReferencingEntity": "contact",
                "ReferencingEntityNavigationPropertyName": "",  # empty nav prop
                "ReferencedEntity": "account",
                "ReferencingAttribute": "parentcustomerid",
                "SchemaName": "contact_customer_accounts",
            }
        ]
        mock_od._entity_set_from_schema_name.return_value = "accounts"

        result = await async_client.query.odata_expands("contact")
        assert result == []

