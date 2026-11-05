# TDS to Report

Import tables from one or more completed TDS Word files into an original Word report. Clauses absent from all uploaded TDS files can be replaced by a user-provided blank-table template so old report values are removed.

## Features

- DOC and DOCX support through a separate LibreOffice conversion container
- Multiple TDS uploads, maximum configurable with `MAX_TDS_FILES`
- Last uploaded TDS wins when the same clause appears in multiple files
- Blank-template fallback clears old report data
- Preserves fonts, sizes, bold/italic, table borders, shading, widths, and merged cells
- Special matching for Clause 8.9.4 `INSULATION DIAGRAM`
- Automatic download using the original report filename plus `_Imported_from_TDS`

## Quick install

```bash
git clone https://github.com/YOUR_GITHUB_NAME/tds-to-report.git
cd tds-to-report
./scripts/install.sh
```

Edit `.env` and set:

```env
APP_IMAGE=ghcr.io/YOUR_GITHUB_NAME/tds-to-report:latest
```

Add private user files locally:

```text
blank-templates/your_blank_tables.docx
standards/your_standard_config.json
```

Start:

```bash
docker compose pull
docker compose up -d
```

Open `http://SERVER_IP:8101`.

## Required local files

This public repository intentionally excludes user templates and standards configuration. See:

- `blank-templates/README.md`
- `standards/README.md`

## Health checks

```bash
curl http://127.0.0.1:8101/api/health
docker compose ps
```

## Local development build

```bash
docker compose -f docker-compose.yml -f docker-compose.build.yml up -d --build
```

## Publish to GHCR

The workflow `.github/workflows/publish-image.yml` builds and publishes `ghcr.io/OWNER/REPOSITORY`. After the first successful workflow, set the package visibility to Public in GitHub Packages so users can pull anonymously.

Create a version release:

```bash
git tag v1.0.0
git push origin v1.0.0
```

## Security

- Never commit `.env`, real reports, TDS files, blank company templates, internal standards JSON, signatures, or customer information.
- Generate the converter token with `openssl rand -hex 32`.
- The LibreOffice converter is reachable only on the internal Docker network.
