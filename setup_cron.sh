#!/usr/bin/env bash
# Sets up a daily cron job to run the NAV tracker at 7:30 AM Pacific Time.
#
# Usage: bash setup_cron.sh
#
# NOTE: Cron uses the system timezone. This script converts 7:30 AM PT to
# your local timezone automatically. If your system is already set to
# America/Los_Angeles or America/Vancouver, it runs at 7:30 directly.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$(command -v python3 || command -v python)"
CRON_LOG="${SCRIPT_DIR}/output/cron.log"

if [ -z "$PYTHON" ]; then
    echo "ERROR: python3 not found. Install Python 3 first."
    exit 1
fi

# Detect system timezone and calculate the right cron hour
SYS_TZ=$(timedatectl show -p Timezone --value 2>/dev/null || cat /etc/timezone 2>/dev/null || echo "unknown")

case "$SYS_TZ" in
    America/Los_Angeles|America/Vancouver|US/Pacific|PST8PDT)
        CRON_HOUR=7
        CRON_MIN=30
        echo "System timezone is Pacific ($SYS_TZ). Scheduling at 7:30 AM local."
        ;;
    America/Denver|America/Edmonton|US/Mountain|MST7MDT)
        CRON_HOUR=8
        CRON_MIN=30
        echo "System timezone is Mountain ($SYS_TZ). Scheduling at 8:30 AM local (= 7:30 AM PT)."
        ;;
    America/Chicago|America/Winnipeg|US/Central|CST6CDT)
        CRON_HOUR=9
        CRON_MIN=30
        echo "System timezone is Central ($SYS_TZ). Scheduling at 9:30 AM local (= 7:30 AM PT)."
        ;;
    America/New_York|America/Toronto|US/Eastern|EST5EDT)
        CRON_HOUR=10
        CRON_MIN=30
        echo "System timezone is Eastern ($SYS_TZ). Scheduling at 10:30 AM local (= 7:30 AM PT)."
        ;;
    *)
        echo "WARNING: Could not detect timezone ($SYS_TZ). Defaulting to 7:30 (assuming Pacific)."
        echo "         Edit your crontab manually if needed: crontab -e"
        CRON_HOUR=7
        CRON_MIN=30
        ;;
esac

CRON_LINE="${CRON_MIN} ${CRON_HOUR} * * 1-5 cd ${SCRIPT_DIR} && ${PYTHON} nav_tracker.py >> ${CRON_LOG} 2>&1"

# Check if cron job already exists
EXISTING=$(crontab -l 2>/dev/null || true)
if echo "$EXISTING" | grep -qF "nav_tracker.py"; then
    echo ""
    echo "A nav_tracker cron job already exists:"
    echo "$EXISTING" | grep "nav_tracker.py"
    echo ""
    read -rp "Replace it? [y/N] " answer
    if [[ "$answer" =~ ^[Yy] ]]; then
        EXISTING=$(echo "$EXISTING" | grep -vF "nav_tracker.py")
    else
        echo "Keeping existing cron job. No changes made."
        exit 0
    fi
fi

# Install cron job
echo "$EXISTING" | { cat; echo "$CRON_LINE"; } | crontab -

echo ""
echo "Cron job installed successfully!"
echo "Schedule: ${CRON_MIN} ${CRON_HOUR} * * 1-5 (Mon-Fri)"
echo "Command:  cd ${SCRIPT_DIR} && ${PYTHON} nav_tracker.py"
echo "Log file: ${CRON_LOG}"
echo ""
echo "Verify with: crontab -l"
