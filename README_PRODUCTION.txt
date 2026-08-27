MODERN TRADE CONTROL TOWER — V63 SUPABASE PRODUCTION

FILES
- app.py: hybrid build. Uses Supabase PostgreSQL when DATABASE_URL is present; falls back to local SQLite when absent.
- migrate_sqlite_to_supabase.py: one-time data migration.
- test_supabase_connection.py: safe connectivity test.
- requirements.txt: production dependencies.

LOCAL STEPS
1. Back up control_tower_data.
2. Keep .env with DATABASE_URL locally; never push .env.
3. pip install -r requirements.txt
4. Run app.py once with DATABASE_URL so PostgreSQL schema is created:
   py -m streamlit run app.py
5. Stop Streamlit after the app opens successfully.
6. Run migration:
   py migrate_sqlite_to_supabase.py
7. Start app again and verify row counts/dashboards.

PERSISTENCE
- Processed data is stored in Supabase PostgreSQL.
- New original upload bytes are also stored in uploads.file_blob, allowing stored source files to be restored after a server restart/redeploy.
- Existing local upload files are copied into file_blob during migration where the stored file still exists.

RENDER DEPLOYMENT
Build command:
  pip install -r requirements.txt
Start command:
  streamlit run app.py --server.address 0.0.0.0 --server.port $PORT
Environment variable:
  DATABASE_URL=<Supabase Session Pooler URI>

SECURITY
Do not commit .env or secrets.toml. Rotate any database password ever shared publicly.


V63.1 POSTGRESQL SQL ALIAS FIX
------------------------------
Fixed PostgreSQL SyntaxError caused by legacy SQLite SQL such as:
    ledger_name AS 'Ledger Name'

PostgreSQL requires:
    ledger_name AS "Ledger Name"

The PostgreSQL compatibility translator now converts legacy AS 'Alias' syntax
before execution, which also protects other exports/dashboards using the same
SQLite alias style.


V63.2 PSYCOPG2 EXECUTION FIX
----------------------------
Fixed:
    IndexError: tuple index out of range
on parameterless PostgreSQL queries containing percent signs such as:
    LIKE '%FG%'

Cause:
The compatibility wrapper always called:
    cursor.execute(sql, ())
Even when the query had no bind parameters. psycopg2 then interpreted %
characters as parameter formatting.

Fix:
- Parameterized query -> cursor.execute(sql, params)
- No parameters       -> cursor.execute(sql)

No database migration is required for V63.2.


V63.3 FAST SUPABASE
===================

Performance-only release. Existing business rules and Supabase data are not changed.

Main fixes:
- Reusable ThreadedConnectionPool (max 6) instead of creating a new remote
  PostgreSQL connection for many individual queries.
- init_db() / schema compatibility runs once per Streamlit server process.
- V58 Ship-to seed runs once per process, not on every Streamlit widget rerun.
- V49 PO mapping schema compatibility runs once per process.
- PostgreSQL startup/index checks run once.
- Existing performance index functions are cached.
- Large source/Main/Factory/B2B/Sales caches use a 5-minute TTL.
- Existing upload/update functions still clear st.cache_data after mutations,
  so newly uploaded information remains visible immediately.

No database migration is required.
Do NOT rerun SQLite-to-Supabase migration.


V63.4 FREE LOW-MEMORY — RENDER FREE STABILITY
==============================================

Reason:
Render Events reported Exit Status 137, consistent with the web process being
killed under memory pressure. V63.3 optimized latency by caching multiple large
pandas DataFrames; on a small Free instance this can retain multiple copies of
Sale Register / Item Ledger / reconciliation data.

Changes:
- PostgreSQL client pool reduced from 6 to 2 connections.
- Removed global cache from cached_table(), preventing persistent copies of
  every full source table.
- Removed full Sale Register dataframe cache from Sales & Return 360.
- Main Reconciliation cache: max 1 entry, 90-second TTL.
- Factory Requirement cache: max 1 entry, 90-second TTL.
- PO/B2B caches capped to one entry.
- Parameterized summary cache capped to 8 entries.
- Cache invalidation now runs Python garbage collection.

No database migration is required.
Supabase data is unchanged.
This build prioritizes Render Free stability over maximum repeated-click speed.
