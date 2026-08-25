# Headway — common tasks. `make setup` once, then `make dev`.

.PHONY: setup dev build test frontend clean

# install JS deps (Tauri CLI + test libs); Rust deps resolve on first build
setup:
	npm install
	@command -v cargo >/dev/null || echo "warning: Rust not found — install it from https://rustup.rs to run/build the desktop app"

# run the desktop app against the live frontend
dev:
	npm run dev

# build the platform installer locally (needs Rust)
build:
	npm run build

# core + headless UI smoke suites
test:
	NODE_PATH=./node_modules node tests/core.test.js
	NODE_PATH=./node_modules node tests/smoke.test.js

# stage the static frontend into dist/ (what the Tauri build bundles)
frontend:
	node scripts/copy-frontend.mjs

clean:
	rm -rf dist src-tauri/target
