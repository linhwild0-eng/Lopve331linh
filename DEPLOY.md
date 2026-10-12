# NTTCLASS deployment

This project includes the web interface and its Express API. Deploy the whole
project as a Node.js service; deploying only `public/index.html` will not provide
shared login or database storage.

## Requirements

- Node.js 22.x
- A hosting service that runs `npm start`
- Persistent disk/storage for SQLite
- One running service instance (sessions and write ordering are held in memory)

## Deploy

1. Upload/extract this project. On Render, set the build command to
   `npm_config_build_from_source=true npm ci` so SQLite is compiled against
   Render's system libraries instead of using a binary that requires a newer
   GLIBC version. On other hosts, install dependencies with `npm ci`; if that
   host has the same GLIBC error, build SQLite from source there as well.
   Keep `public/index.html` at the `public` folder directly under the project
   root. When deploying from GitHub, make sure this file is committed and set
   Render's Root Directory to the folder containing `server.js`, `package.json`,
   and `public`.
2. Configure the service start command as `npm start`. The service must use its
   assigned `PORT`; the app reads that environment variable automatically.
3. Configure `NTTCLASS_DB_PATH` to a file on the host's persistent disk, such as
   `/var/data/nttclass.db`. Ensure the service can write to that disk.
4. Before the first start on an empty database, set `NTTCLASS_ADMIN_USERNAME`,
   `NTTCLASS_ADMIN_PASSWORD`, and optionally `NTTCLASS_ADMIN_DISPLAY_NAME`.
   The password must be at least 8 characters. Do not publish `.env` or put
   credentials in source control. The app will fail to start if these values are
   missing in production.
5. In Render's Environment settings, set `NTTCLASS_ACCOUNT_CREATION_CODE` to a
   private value of at least 12 bytes. Keep it secret and do not put its actual
   value in source control. Creating an account requires both an Admin login
   and this exact code in the account form. If the variable is unset, account
   creation is disabled; a configured code shorter than 12 bytes prevents the
   service from starting.
6. Open the public HTTPS URL provided by the host and sign in with the configured
   administrator account.

For local development only, the app can fall back to `admin` / `ChangeMe123!` if
`NODE_ENV` is not `production`, but this is only for convenience and should not
be used in a shared or internet-facing deployment.

## Updating an existing deployment

- Keep the persistent SQLite database and its storage path. Do not overwrite or
  delete it while replacing application files.
- Back up the database before deploying an update. If moving hosts, transfer the
  database using the hosting provider's private storage/backup tools, not a
  public website or source archive.
- `nttclass.db`, `.env`, and `node_modules` are intentionally not part of the
  source deployment archive. The database contains private school records;
  configure its persistent path separately on the hosting service.
- After deploy, verify `/api/health`, sign in, and confirm existing classes and
  students are still present.

## Local run

```powershell
npm ci
$env:PORT = "3000"
npm start
```

The database defaults to `nttclass.db` next to `server.js`. Set
`NTTCLASS_DB_PATH` to another writable location to use a different database.
