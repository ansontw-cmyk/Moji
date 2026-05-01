#!/bin/bash
set -euo pipefail

# Only run in remote (Claude Code web) environments
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

echo '{"async": true, "asyncTimeout": 300000}'

# ── 1. Install npm dependencies ───────────────────────────────────────────────
if [ -f "$CLAUDE_PROJECT_DIR/package.json" ]; then
  echo "[session-start] Installing npm dependencies..."
  cd "$CLAUDE_PROJECT_DIR"
  npm install
fi

# ── 2. Install agency-agents to ~/.claude/agents/ ────────────────────────────
AGENTS_DIR="$HOME/.claude/agents"

if [ ! -d "$AGENTS_DIR" ] || [ -z "$(ls -A "$AGENTS_DIR" 2>/dev/null)" ]; then
  echo "[session-start] Installing agency-agents..."
  mkdir -p "$AGENTS_DIR"

  REPO_DIR="/tmp/agency-agents"

  # Clone if not already present
  if [ ! -d "$REPO_DIR" ]; then
    git clone --depth=1 https://github.com/msitarzewski/agency-agents "$REPO_DIR"
  fi

  # Run the official install script for Claude Code
  cd "$REPO_DIR"
  bash scripts/install.sh --tool claude-code

  echo "[session-start] Agents installed: $(ls "$AGENTS_DIR" | wc -l) agents"
else
  echo "[session-start] Agents already installed: $(ls "$AGENTS_DIR" | wc -l) agents"
fi
