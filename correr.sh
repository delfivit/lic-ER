#!/bin/bash
cd "$(dirname "$0")"
exec ./.venv/bin/python -u scraper.py "$@"
