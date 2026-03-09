# Missing Thorn CRM & Dashboard System

Private monorepo for the Missing Thorn BI dashboard build pipeline.

## Quick Start

```bash
# 1. Copy .env.example to .env and add your GitHub token
cp .env.example .env

# 2. Build dashboards (point --data to your Excel files)
make build DATA_DIR=~/OneDrive/MT\ Dashboard\ Data/

# 3. Deploy to GitHub Pages
make deploy
```

## Structure

| Directory | Contents |
|-----------|----------|
| `templates/` | HTML dashboard templates (with data injection markers) |
| `scripts/` | Python build scripts that read Excel → inject data |
| `docs/` | Firebase setup, CRM integration notes |
| `dist/` | Built output (index.html, index_rep.html) |

## Dashboards

- **Executive Dashboard** → deployed to [robbied112/mt-dashboard](https://github.com/robbied112/mt-dashboard)
- **Rep Dashboard** → deployed to [robbied112/mt-rep-dashboard](https://github.com/robbied112/mt-rep-dashboard)

## Build Commands

```bash
make build                    # Build both (data from ./data/)
make build DATA_DIR=/path/    # Build both with custom data dir
make build-exec               # Executive dashboard only
make build-rep                # Rep dashboard only
make deploy                   # Deploy both to GitHub Pages
make all DATA_DIR=/path/      # Build + deploy everything
```
