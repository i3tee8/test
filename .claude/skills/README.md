# Installed skills

These skills come from the [UI UX Pro Max](https://github.com/nextlevelbuilder/ui-ux-pro-max-skill)
project (MIT, © Next Level Builder).

- **Version:** 2.13.0
- **Upstream commit:** f23267105ad1f4ccd94af45d382584ad45b586f7
- **Installed:** 2026-09-01

## Contents

| Skill | Purpose |
|-------|---------|
| `ui-ux-pro-max` | Orchestrator: searchable UI/UX database (styles, palettes, typography, UX guidelines, charts, per-stack guidance) plus the design-system generator |
| `design` | Brand identity, design tokens, logo/icon/CIP generation, slides |
| `design-system` | Three-layer token architecture and component specs |
| `brand` | Brand voice, visual identity, messaging frameworks |
| `ui-styling` | shadcn/ui + Tailwind component styling |
| `slides` | Strategic HTML presentations |
| `banner-design` | Social, ad, hero, and print banners |

## Requirements

Python 3.x (standard library only — the scripts install nothing and make no
network calls).

## Using the search tool directly

Run from the project root:

```bash
python3 ".claude/skills/ui-ux-pro-max/scripts/search.py" "minimalism dark mode" --domain style
python3 ".claude/skills/ui-ux-pro-max/scripts/search.py" "saas analytics dashboard" --design-system -p "My Project"
python3 ".claude/skills/ui-ux-pro-max/scripts/search.py" "streaming suspense" --stack nextjs
```

## Updating

Re-copy `.claude/skills/` from a fresh checkout of the upstream repo, then
re-apply the one local change: `ui-ux-pro-max/SKILL.md` uses project-relative
script paths (`.claude/skills/ui-ux-pro-max/scripts/search.py`) instead of
upstream's `${CLAUDE_PLUGIN_ROOT}/...`, which is only defined when the skills
are loaded as a Claude Code plugin rather than as project skills.

Alternatively, install upstream's way with their CLI (`npx ui-ux-pro-max-cli
init --ai claude`) or as a plugin:

```
/plugin marketplace add nextlevelbuilder/ui-ux-pro-max-skill
/plugin install ui-ux-pro-max@ui-ux-pro-max-skill
```
