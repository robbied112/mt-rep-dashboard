#!/bin/bash
# Deploy built dashboards to GitHub Pages repos.
#
# Prerequisites:
#   - Set GITHUB_TOKEN env var (or create .env file with GITHUB_TOKEN=ghp_xxx)
#   - dist/index.html and dist/index_rep.html must exist (run make build first)
#
# Usage:
#   ./deploy.sh          # Deploy both dashboards
#   ./deploy.sh exec     # Deploy executive dashboard only
#   ./deploy.sh rep      # Deploy rep dashboard only

set -euo pipefail

# Load .env if it exists
if [ -f .env ]; then
    export $(grep -v '^#' .env | xargs)
fi

if [ -z "${GITHUB_TOKEN:-}" ]; then
    echo "ERROR: GITHUB_TOKEN not set."
    echo "Set it via: export GITHUB_TOKEN=ghp_xxx"
    echo "Or create a .env file with: GITHUB_TOKEN=ghp_xxx"
    exit 1
fi

GITHUB_USER="robbied112"
EXEC_REPO="mt-dashboard"
REP_REPO="mt-rep-dashboard"

DEPLOY_TARGET="${1:-all}"
TMPDIR=$(mktemp -d)
trap "rm -rf $TMPDIR" EXIT

deploy_repo() {
    local source_file="$1"
    local repo_name="$2"
    local label="$3"

    if [ ! -f "$source_file" ]; then
        echo "ERROR: $source_file not found. Run 'make build' first."
        return 1
    fi

    echo "Deploying $label to $GITHUB_USER/$repo_name..."

    local repo_dir="$TMPDIR/$repo_name"
    git clone --depth 1 "https://${GITHUB_TOKEN}@github.com/${GITHUB_USER}/${repo_name}.git" "$repo_dir" 2>/dev/null

    cp "$source_file" "$repo_dir/index.html"

    cd "$repo_dir"
    if git diff --quiet; then
        echo "  No changes to deploy for $label."
    else
        git add index.html
        git commit -m "Update dashboard $(date +%Y-%m-%d)"
        git push
        echo "  Deployed $label successfully."
    fi
    cd - > /dev/null
}

case "$DEPLOY_TARGET" in
    exec)
        deploy_repo "dist/index.html" "$EXEC_REPO" "Executive Dashboard"
        ;;
    rep)
        deploy_repo "dist/index_rep.html" "$REP_REPO" "Rep Dashboard"
        ;;
    all)
        deploy_repo "dist/index.html" "$EXEC_REPO" "Executive Dashboard"
        deploy_repo "dist/index_rep.html" "$REP_REPO" "Rep Dashboard"
        ;;
    *)
        echo "Usage: ./deploy.sh [exec|rep|all]"
        exit 1
        ;;
esac

echo "Done."
