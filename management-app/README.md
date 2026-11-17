# Catalogue management app

The app is a static, read-only view over the repository's generated
`outputs/reports/PICKLIST_SYNC.csv`, `WORKING_LIST_RECONCILIATION.csv`, and
`PRODUCTION_READINESS.json`, plus `data/picklist_assignments.json`. PLUs stay
strings. Approval, identity, quality, rights, provenance, reuse and release
status are shown separately.

## Run locally

From the repository root, serve the repository over HTTP so the browser can
load its report and image assets:

```bash
python3 -m http.server 8000
```

Open `http://localhost:8000/management-app/`. The page requires the current
generated reports and assignment registry; it displays a visible error instead
of substituting sample data if any source is missing or invalid.

Run the app-model regression tests with Node.js 20 or later:

```bash
node --test management-app/app-model.test.mjs
```

## Static deployment

GitHub Pages can serve the repository root as static files; the app URL is
`https://<owner>.github.io/<repository>/management-app/`. Ensure the generated
reports are present in the published site and keep the relative repository
paths intact. The app makes no write requests and contains no credentials.

## Write authorization blocker

No authenticated server-side API or approval service is configured in this
repository. GitHub Pages cannot safely authorize catalogue edits, image
approvals or release publication. The app therefore cannot submit decisions or
approve/publish anything. Enabling writes requires a server-side service that
authenticates authorized reviewers, validates against the authoritative
PickList/WorkingList identities and release policy, records an auditable
decision, and uses a least-privilege GitHub App installation token on the
server. Never place privileged tokens in browser code or local storage.
