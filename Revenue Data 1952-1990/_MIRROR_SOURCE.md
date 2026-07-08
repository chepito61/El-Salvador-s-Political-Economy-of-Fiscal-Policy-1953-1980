# Mirror provenance

This folder is a mirror of the co-authored revenue dataset repository:

- **Source repo:** https://github.com/lucaschavesgurgel/EL-SALVADOR
- **Source commit:** `423b07ac897b91d590add44eb6043c22c929ba72`
- **Source commit date:** 2026-03-26
- **Mirrored on:** 2026-03 (via GitHub REST API)

The original `README.md` from the source repo is preserved unchanged in this folder.

**Not mirrored** (local IDE / tooling state only): `.Rhistory`, `.claude/`, `.vscode/`, `.gitattributes`.

## Known data-quality notes (flagged, not yet corrected in source)
1. **1964 coffee tax** ("Sobre el Café") appears under `item_code 71` with value ₡30,461,430.31 — identical to the 1954 value under `item_code 21`. Likely a duplication/copy error in extraction; verify against MH source before use.
2. **1954 `Total General de Ingresos`** is blank; only `Total Ingresos Ordinarios` (₡167.4M) is present.
3. **1952 revenue definition** differs from the 1952–53 notes in the main repo (source: ₡126.09M ordinary / ₡129.69M general-fund perceived; notes: ₡135.16M percibidas y devengadas). Reconcile accrual-vs-cash and ordinary-vs-general-fund scope before combining datasets.
