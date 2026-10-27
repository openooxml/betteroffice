# WOPI smoke compose files

Throwaway local stacks for manually or scripted smoke testing this app
against a real WOPI host. They use HTTP on one Docker host and demo
credentials from the environment. Never use them in production, and never
commit real credentials: both files fail to start until you set the admin
password variable.

Add the smoke hostnames to `/etc/hosts` on the Docker host first:

```text
127.0.0.1 cloud.opencloud.test office.opencloud.test
127.0.0.1 cloud.nextcloud.test office.nextcloud.test
```

## OpenCloud

```bash
export OC_ADMIN_PASSWORD='choose-a-throwaway-password'
docker compose -f apps/wopi/deploy/opencloud.compose.yml up --build
```

Open `http://cloud.opencloud.test:9200`, sign in as `admin` with
`OC_ADMIN_PASSWORD` (demo users are created by default; set
`OC_DEMO_USERS=false` to skip them), upload or create a DOCX, XLSX or PPTX
file, and open it. It should load in BetterOffice, save edits back, and let
a second user join the same file.

Check the app alone with `curl http://office.opencloud.test:3000/healthz`.

## Nextcloud

```bash
export NC_ADMIN_PASSWORD='choose-a-throwaway-password'
docker compose -f apps/wopi/deploy/nextcloud.compose.yml up --build
```

Then install the Nextcloud Office app and point it at this app:

```bash
docker compose -f apps/wopi/deploy/nextcloud.compose.yml exec nextcloud php occ app:install richdocuments
docker compose -f apps/wopi/deploy/nextcloud.compose.yml exec nextcloud php occ config:app:set richdocuments wopi_url --value='http://office.nextcloud.test:3000'
docker compose -f apps/wopi/deploy/nextcloud.compose.yml exec nextcloud php occ config:app:set richdocuments public_wopi_url --value='http://office.nextcloud.test:3000'
docker compose -f apps/wopi/deploy/nextcloud.compose.yml exec nextcloud php occ config:system:set allow_local_remote_servers --value=true --type=boolean
```

In Settings, Administration, Nextcloud Office, add the egress address the
Nextcloud container sees for the `wopi` service to the allow list for WOPI
requests. For this local stack, allow the Docker bridge range of the
compose network (find it with `docker network inspect`), or, for a quick
local check only, the address of the `wopi` container. Then open
`http://cloud.nextcloud.test:8080`, sign in as the admin from
`NC_ADMIN_USER` and `NC_ADMIN_PASSWORD`, and open a DOCX, XLSX or PPTX
file.

Check the app alone with `curl http://office.nextcloud.test:3000/healthz`.

## Cleanup

```bash
docker compose -f apps/wopi/deploy/opencloud.compose.yml down -v
docker compose -f apps/wopi/deploy/nextcloud.compose.yml down -v
```
