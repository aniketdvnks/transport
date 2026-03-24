#!/usr/bin/env bash

set -euo pipefail

SITE="${SITE:-transport}"
BRANCH="${BRANCH:-version-15}"
REPORT_DATE="${REPORT_DATE:-2026-03-24}"
SKIP_BACKUP=0
SEED_HAZCHEM_DATA=0

usage() {
	cat <<'EOF'
Usage:
  ./scripts/sync_remote_version_15.sh [options]

Options:
  --site <site>              Frappe site name. Default: transport
  --branch <branch>          Git branch to sync. Default: version-15
  --report-date <date>       Report verification date. Default: 2026-03-24
  --skip-backup              Skip bench backup before syncing
  --seed-hazchem-data        Seed the hazchem demo/operational dataset after migrate
  -h, --help                 Show this help

Examples:
  ./scripts/sync_remote_version_15.sh
  ./scripts/sync_remote_version_15.sh --site transport --branch version-15
  ./scripts/sync_remote_version_15.sh --seed-hazchem-data --report-date 2026-03-24
EOF
}

log() {
	printf '\n[%s] %s\n' "$1" "$2"
}

run() {
	log "RUN" "$*"
	"$@"
}

while [[ $# -gt 0 ]]; do
	case "$1" in
		--site)
			SITE="$2"
			shift 2
			;;
		--branch)
			BRANCH="$2"
			shift 2
			;;
		--report-date)
			REPORT_DATE="$2"
			shift 2
			;;
		--skip-backup)
			SKIP_BACKUP=1
			shift
			;;
		--seed-hazchem-data)
			SEED_HAZCHEM_DATA=1
			shift
			;;
		-h|--help)
			usage
			exit 0
			;;
		*)
			echo "Unknown argument: $1" >&2
			usage
			exit 1
			;;
	esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
BENCH_ROOT="$(cd "${APP_ROOT}/../.." && pwd)"

if [[ ! -d "${BENCH_ROOT}/sites" ]]; then
	echo "Could not determine bench root from ${SCRIPT_DIR}" >&2
	exit 1
fi

cd "${BENCH_ROOT}"

if [[ ! -d "${APP_ROOT}/.git" ]]; then
	echo "App repository not found at ${APP_ROOT}" >&2
	exit 1
fi

if ! command -v bench >/dev/null 2>&1; then
	echo "bench command not found in PATH" >&2
	exit 1
fi

if [[ "${SKIP_BACKUP}" -eq 0 ]]; then
	run bench --site "${SITE}" backup --with-files
else
	log "INFO" "Skipping backup because --skip-backup was provided"
fi

cd "${APP_ROOT}"

if [[ -n "$(git status --porcelain)" ]]; then
	STASH_NAME="pre-${BRANCH}-sync-$(date +%F-%H%M%S)"
	run git stash push -u -m "${STASH_NAME}"
	log "INFO" "Local changes were stashed as ${STASH_NAME}"
else
	log "INFO" "Git worktree is clean"
fi

run git fetch origin --prune

if git show-ref --verify --quiet "refs/heads/${BRANCH}"; then
	run git checkout "${BRANCH}"
else
	run git checkout -b "${BRANCH}" "origin/${BRANCH}"
fi

run git branch --set-upstream-to="origin/${BRANCH}" "${BRANCH}"
run git pull --ff-only origin "${BRANCH}"

cd "${BENCH_ROOT}"

run bench --site "${SITE}" migrate
run bench build --app trans_ms
run bench --site "${SITE}" clear-cache
run bench --site "${SITE}" clear-website-cache

if [[ "${SEED_HAZCHEM_DATA}" -eq 1 ]]; then
	run bench --site "${SITE}" execute trans_ms.transport_management.hazchem_operations.setup_hazchem_operations --kwargs "{'report_date': '${REPORT_DATE}'}"
else
	log "INFO" "Skipping hazchem seed data. Pass --seed-hazchem-data to load the sample dataset."
fi

log "RUN" "Verifying Daily Trip Schedule"
bench --site "${SITE}" execute frappe.desk.query_report.run --kwargs "{'report_name': 'Daily Trip Schedule', 'filters': {'report_date': '${REPORT_DATE}'}}"

log "RUN" "Verifying Vehicle Tracking Report"
bench --site "${SITE}" execute frappe.desk.query_report.run --kwargs "{'report_name': 'Vehicle Tracking Report', 'filters': {'from_date': '${REPORT_DATE}', 'to_date': '${REPORT_DATE}'}}"

if bench restart; then
	log "INFO" "bench restart completed"
else
	log "WARN" "bench restart failed. Restart the bench manually if your environment requires it."
fi

log "DONE" "Remote sync completed for site ${SITE} on branch ${BRANCH}"
