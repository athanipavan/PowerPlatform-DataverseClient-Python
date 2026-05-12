# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

"""Async query operations namespace for the Dataverse SDK."""

from __future__ import annotations

from typing import Any, Dict, List, TYPE_CHECKING

from ...core.errors import MetadataError
from ...models.record import Record

if TYPE_CHECKING:
    from ..async_client import AsyncDataverseClient


__all__ = ["AsyncQueryOperations"]


class AsyncQueryOperations:
    """Async namespace for query operations.

    Accessed via ``client.query``. Provides query and search operations
    against Dataverse tables.

    :param client: The parent :class:`~PowerPlatform.Dataverse.aio.async_client.AsyncDataverseClient` instance.
    :type client: ~PowerPlatform.Dataverse.aio.async_client.AsyncDataverseClient

    Example::

        async with AsyncDataverseClient(base_url, credential) as client:

            # Fluent query builder (recommended)
            from PowerPlatform.Dataverse.models.filters import col

            for record in await (client.query.builder("account")
                                 .select("name", "revenue")
                                 .where(col("statecode") == 0)
                                 .order_by("revenue", descending=True)
                                 .top(100)
                                 .execute()):
                print(record["name"])

            # SQL query
            rows = await client.query.sql("SELECT TOP 10 name FROM account ORDER BY name")
            for row in rows:
                print(row["name"])
    """

    def __init__(self, client: "AsyncDataverseClient") -> None:
        self._client = client

    # -------------------------------------------------------------------- sql

    async def sql(self, sql: str) -> List[Record]:
        """Execute a read-only SQL query using the Dataverse Web API.

        The Dataverse SQL endpoint supports a broad subset of T-SQL::

            SELECT / SELECT DISTINCT / SELECT TOP N (0-5000)
            FROM table [alias]
            INNER JOIN / LEFT JOIN (multi-table, no depth limit)
            WHERE (=, !=, >, <, >=, <=, LIKE, IN, NOT IN, IS NULL,
                   IS NOT NULL, BETWEEN, AND, OR, nested parentheses)
            GROUP BY column
            ORDER BY column [ASC|DESC]
            OFFSET n ROWS FETCH NEXT m ROWS ONLY
            COUNT(*), SUM(), AVG(), MIN(), MAX()

        ``SELECT *`` is not supported -- specify column names explicitly.
        Use :meth:`sql_columns` to discover available column names for a table.

        Not supported: SELECT *, subqueries, CTE, HAVING, UNION,
        RIGHT/FULL/CROSS JOIN, CASE, COALESCE, window functions,
        string/date/math functions, INSERT/UPDATE/DELETE. For writes, use
        ``client.records`` methods.

        :param sql: Supported SQL SELECT statement.
        :type sql: :class:`str`

        :return: List of :class:`~PowerPlatform.Dataverse.models.record.Record`
            objects. Returns an empty list when no rows match.
        :rtype: list[~PowerPlatform.Dataverse.models.record.Record]

        :raises ~PowerPlatform.Dataverse.core.errors.ValidationError:
            If ``sql`` is not a string or is empty.

        Example:
            Basic query::

                rows = await client.query.sql(
                    "SELECT TOP 10 name FROM account ORDER BY name"
                )

            JOIN with aggregation::

                rows = await client.query.sql(
                    "SELECT a.name, COUNT(c.contactid) as cnt "
                    "FROM account a "
                    "JOIN contact c ON a.accountid = c.parentcustomerid "
                    "GROUP BY a.name"
                )
        """
        async with self._client._scoped_odata() as od:
            rows = await od._query_sql(sql)
            return [Record.from_api_response("", row) for row in rows]

    # --------------------------------------------------------------- sql_columns

    async def sql_columns(
        self,
        table: str,
        *,
        include_system: bool = False,
    ) -> List[Dict[str, Any]]:
        """Return a simplified list of SQL-usable columns for a table.

        Each dict contains ``name`` (logical name for SQL), ``type``
        (Dataverse attribute type), ``is_pk`` (primary key flag), and
        ``label`` (display name).  Virtual columns are always excluded
        because the SQL endpoint cannot query them.

        :param table: Schema name of the table (e.g. ``"account"``).
        :type table: :class:`str`
        :param include_system: When ``False`` (default), columns that end
            with common system suffixes (``_base``, ``versionnumber``,
            ``timezoneruleversionnumber``, ``utcconversiontimezonecode``,
            ``importsequencenumber``, ``overriddencreatedon``) are excluded.
        :type include_system: :class:`bool`

        :return: List of column metadata dicts.
        :rtype: list[dict[str, typing.Any]]

        Example::

            cols = await client.query.sql_columns("account")
            for c in cols:
                print(f"{c['name']:30s} {c['type']:20s} PK={c['is_pk']}")
        """
        _SYSTEM_SUFFIXES = (
            "_base",
            "versionnumber",
            "timezoneruleversionnumber",
            "utcconversiontimezonecode",
            "importsequencenumber",
            "overriddencreatedon",
        )

        raw = await self._client.tables.list_columns(
            table,
            select=[
                "LogicalName",
                "SchemaName",
                "AttributeType",
                "IsPrimaryId",
                "IsPrimaryName",
                "DisplayName",
                "AttributeOf",
            ],
            filter="AttributeType ne 'Virtual'",
        )
        result: List[Dict[str, Any]] = []
        for c in raw:
            name = c.get("LogicalName", "")
            if not name:
                continue
            if not include_system and any(name.endswith(s) for s in _SYSTEM_SUFFIXES):
                continue
            # Skip computed display-name columns (AttributeOf is set, meaning
            # they are auto-generated from a lookup column)
            if c.get("AttributeOf"):
                continue
            # Extract display label
            label = ""
            dn = c.get("DisplayName")
            if isinstance(dn, dict):
                ul = dn.get("UserLocalizedLabel")
                if isinstance(ul, dict):
                    label = ul.get("Label", "")
            result.append(
                {
                    "name": name,
                    "type": c.get("AttributeType", ""),
                    "is_pk": bool(c.get("IsPrimaryId")),
                    "is_name": bool(c.get("IsPrimaryName")),
                    "label": label,
                }
            )
        result.sort(key=lambda x: (not x["is_pk"], not x["is_name"], x["name"]))
        return result

    # =========================================================================
    # OData helpers -- discover columns, navigation properties, and bind values
    # =========================================================================

    # ------------------------------------------------------- odata_expands

    async def odata_expands(
        self,
        table: str,
    ) -> List[Dict[str, Any]]:
        """Discover all ``$expand`` navigation properties from a table.

        Returns entries for each outgoing lookup (single-valued navigation
        property).  Each entry contains the exact PascalCase navigation
        property name needed for ``$expand`` and ``@odata.bind``, plus
        the target entity set name.

        :param table: Schema name of the table (e.g. ``"contact"``).
        :type table: :class:`str`

        :return: List of dicts, each with:

            - ``nav_property`` -- PascalCase navigation property for $expand
            - ``target_table`` -- target entity logical name
            - ``target_entity_set`` -- target entity set (for @odata.bind)
            - ``lookup_attribute`` -- the lookup column logical name
            - ``relationship`` -- relationship schema name

        :rtype: list[dict[str, typing.Any]]

        Example::

            expands = await client.query.odata_expands("contact")
            for e in expands:
                print(f"expand={e['nav_property']}  -> {e['target_table']}")

            # Use in a query
            e = next(e for e in expands if e['target_table'] == 'account')
            records = await client.records.list("contact",
                                               select=["fullname"],
                                               expand=[e['nav_property']])
        """
        table_lower = table.lower()
        rels = await self._client.tables.list_table_relationships(table)

        result: List[Dict[str, Any]] = []
        for r in rels:
            ref_entity = (r.get("ReferencingEntity") or "").lower()
            if ref_entity != table_lower:
                continue
            nav_prop = r.get("ReferencingEntityNavigationPropertyName", "")
            target = r.get("ReferencedEntity", "")
            lookup_attr = r.get("ReferencingAttribute", "")
            schema = r.get("SchemaName", "")
            if not nav_prop or not target:
                continue

            # Resolve entity set name for @odata.bind
            target_set = ""
            try:
                async with self._client._scoped_odata() as od:
                    target_set = await od._entity_set_from_schema_name(target)
            except (KeyError, AttributeError, ValueError, MetadataError):
                pass  # Entity set resolution failed; target_set stays empty

            result.append(
                {
                    "nav_property": nav_prop,
                    "target_table": target,
                    "target_entity_set": target_set,
                    "lookup_attribute": lookup_attr,
                    "relationship": schema,
                }
            )

        result.sort(key=lambda x: (x["target_table"], x["nav_property"]))
        return result
