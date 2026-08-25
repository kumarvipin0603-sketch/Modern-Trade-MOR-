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
