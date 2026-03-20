# LEARNINGS.md

| Date | Component | Issue | Resolution | Insight |
|------|-----------|-------|------------|---------|
| 2026-03-20 | Project Setup | Initial scaffold | Created project structure per spec | Follow numbered-section pattern in app.py for maintainability |
| 2026-03-20 | Slab Colors | Tables lacked visual slab distinction | Added color_light field to SLAB_CONFIG; used SLAB_COLORS_LIGHT for row backgrounds in Summary and Dealer Details tables | Always add new SLAB_CONFIG fields to both app.py and scripts/generate_slab_report.py, update blueprint docs and test_helpers required-keys check |
| 2026-03-20 | Dealer Details | Needed bifurcated qualified volumes and next-slab info | Replaced Qualified Volume with Qual. Shop Vol. + Qual. Site Vol.; added Next Slab and Volume to Qualify columns | Use display-level column renaming (not modifying source DataFrame) to present internal column names differently in UI; mirror new derived columns in both app.py and generate_slab_report.py |
