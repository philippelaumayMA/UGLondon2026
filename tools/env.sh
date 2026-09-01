# Source before any deck authoring command:  source tools/env.sh
#
#   SKILL                     the moodys-pptx skill, whose scripts/ holds ml and md
#   MOODYS_ASSETS_DIR         where md.ensure_icons() caches the official icon set.
#                             Required because CLAUDE_PLUGIN_DATA is unset in a plain
#                             shell; without it every icon is a blue circle.
#   DYLD_FALLBACK_LIBRARY_PATH  macOS only. cairocffi dlopens libcairo by bare name
#                             and does not search Homebrew's prefix, so `import
#                             cairosvg` raises OSError and icon rendering dies. Needs
#                             `brew install cairo`.
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
export DYLD_FALLBACK_LIBRARY_PATH=/opt/homebrew/lib:/usr/local/lib:/usr/lib
