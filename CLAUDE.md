# Fabrication Database — Claude Context

## What this project is
A Windows tool for managing lab samples and fabrication process data during Benjamin Haubmann's PhD at TUM. It tracks experimental samples with 16-character IDs, links process results to samples, syncs folders, and logs to Excel. Available as both a CLI (`base.py`) and a PyQt5 GUI (`nosferatool.py`) that share the same pickle database and core logic — the GUI reuses `base.py` functions directly (e.g. `core.tag()`, `core.cleave()`) rather than reimplementing them, so the two stay interchangeable.

## Key facts
- **Core logic**: `base.py` (CLI + shared functions, ~1150 lines)
- **GUI**: `nosferatool.py` (PyQt5 front-end, formerly `base_gui.py`) + Qt Designer layouts `base_gui.ui`, `add_create_dialog.ui`, `manage_entry_dialog.ui`, theme `dark_orange.xml`
- **Platform**: Windows-only (Win32 COM, `I:` drive paths, pickle DB)
- **Database**: Python pickle at `I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\IDbase`
- **Excel**: `Sample Overview Local.xlsx` on NAS; shared version at `C:\Users\ge23jud\OneDrive - TUM\Sample Overview.xlsx`. The `_get_excel_row()` helper tries multiple paths with fallback.
- **GitHub**: https://github.com/ge23jud/Fabrication-Database.git
- **Active branch**: `master`
- **Desktop shortcut**: `Nosferatool.lnk` launches `pythonw nosferatool.py`. Note: editing an existing `.lnk` via `WScript.Shell`'s `CreateShortcut` and setting only one property (then `.Save()`) blanks out all other properties — always set TargetPath/Arguments/WorkingDirectory/IconLocation together when updating a shortcut.

## ID format
16-character IDs: `YYYYMMDD` + type prefix + number (e.g. `20240715epi1780`, `20240715spl2407`).
Process types: `sem`, `plm`, `epi`, `elx`, `mic`, `xrd`, `tem`, `mla`, `rie`, `dek`.
Special (excluded from tagging): `spl` (sample), `des`, `sim`, `scr`, `ana`.

## Excel structure (Tabelle1)
Rows = samples in sorted DB order; row = sample_index + 2. Columns are always resolved by header name, never a hardcoded letter.
Full column list: A=Name, B=Type, C=Cleaved From, D=eSAE, E=Ellipsometry, F=Clean, G=Spin-Coating, H=Elionix, I=MLA, J=Development, K=Evaporator, L=Lift-Off, M=RIE, N=HF 1, O=HF 2, P=Growth, Q=NW Transfer, R=Pick and Place, S=Optical Microscope, T=Dektak, U=SEM, V=PL, W=XRD, X=TEM, Y=Design, Z=Cleaved, AA=NWs removed by Pick and Place, AB=Current Location.
Note: "Cleaved From"/"Cleaved" here are Excel-only, read-only growth-tracing fields — a separate concept from the DB-native `cleave_parents`/`cleave_children` below (deliberately different naming to avoid confusion).

## Commands (python base.py <command>)
`add`, `goto`, `ls`, `delete`, `checkall`, `update`, `display`, `update_readme`, `comment`, `inspect`, `edit_readme`, `create`, `new_sample`, `tag`, `untag`, `sync`, `sync_all`, `tags`, `untagged`, `info`

(`cleave()` is a shared function, not a standalone subcommand — it's called internally by `new_sample()` when a "cleaved from" sample is given.)

## Sample cleave relationships (`cleave_parents` / `cleave_children`)
DB-native, bidirectional link recorded when a sample is physically cleaved from another. `cleave(ID, parent_name, param="")` resolves `parent_name` against sample keys (same substring convention as `tag()`), then appends `{"sample": ..., "param": ...}` to the child's `cleave_parents` and the parent's `cleave_children`. `param` is free text (e.g. pitch/dose/diameter or which piece — "SEM", "Transfer p8"). Both `new_sample()` (CLI, interactive prompt) and the GUI's New Sample forms (`newSampleCleavedFromEdit`/`newSampleParamEdit` fields) call this same function. Dedup-safe to call repeatedly (checks for an existing matching parent entry first).

## `info` command
`python base.py info spl2407` or `python base.py info epi1780`
- **Sample mode**: DB info/comments → tagged processes → cleave relationships (`Cleaved from:`/`Cleaved into:`, DB-native, only shown if non-empty) → Sample Overview (filtered columns) → growth origin
- **Process mode**: DB info/comments → tagged samples (name + status only, no filepath) — no Sample Overview section for process-type IDs
- **Growth origin** (sample mode, Excel-derived): checks P (growth wafer itself), Q with "from" keyword (NW transfer), C (cleaved from parent → looks up parent's P column)
- Multiline Excel cells joined with ` | `

## NAS path
`\\nas.ads.mwn.de\tuze\wsi\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\`
The `I:` drive maps to this location.
