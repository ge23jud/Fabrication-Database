# Fabrication Database — Claude Context

## What this project is
A Windows CLI tool (`base.py`) for managing lab samples and fabrication process data during Benjamin Haubmann's PhD at TUM. It tracks experimental samples with 16-character IDs, links process results to samples, syncs folders, and logs to Excel.

## Key facts
- **Single-file codebase**: `base.py` (~1100 lines)
- **Platform**: Windows-only (Win32 COM, `I:` drive paths, pickle DB)
- **Database**: Python pickle at `I:\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\IDbase`
- **Excel**: `Sample Overview Local.xlsx` on NAS; shared version at `C:\Users\ge23jud\OneDrive - TUM\Sample Overview.xlsx`. The `_get_excel_row()` helper tries multiple paths with fallback.
- **GitHub**: https://github.com/ge23jud/Fabrication-Database.git
- **Active branch**: `claude/access-fabrication-database-ha3IT`

## ID format
16-character IDs: `YYYYMMDD` + type prefix + number (e.g. `20240715epi1780`, `20240715spl2407`).
Process types: `sem`, `plm`, `epi`, `elx`, `mic`, `xrd`, `tem`, `mla`, `rie`.
Special (excluded from tagging): `spl` (sample), `des`, `sim`, `scr`, `ana`.

## Excel structure (Tabelle1)
Rows = samples in sorted DB order; row = sample_index + 2.
Key columns: A=Name, B=Type, C=Cleaved From, D=eSAE, E=Ellipsometry, F=Clean, G=Spin-Coating, H=Elionix, I=MLA, J=Development, M=RIE, N=HF 1, O=HF 2, P=Growth, Q=NW Transfer, X=Design, Y=Cleaved.

## Commands (python base.py <command>)
`add`, `goto`, `ls`, `delete`, `checkall`, `update`, `display`, `update_readme`, `comment`, `inspect`, `edit_readme`, `create`, `new_sample`, `tag`, `untag`, `sync`, `sync_all`, `tags`, `untagged`, `info`

## Recently added: `info` command
`python base.py info spl2407` or `python base.py info epi1780`
- **Sample mode**: DB info/comments → tagged processes → Sample Overview (filtered columns) → growth origin
- **Process mode**: DB info/comments → tags → relevant Excel cell per tagged sample
- **Growth origin**: checks P (growth wafer itself), Q with "from" keyword (NW transfer), C (cleaved from parent → looks up parent's P column)
- Multiline Excel cells joined with ` | `

## NAS path
`\\nas.ads.mwn.de\tuze\wsi\e24\SQN\Researchers\Haubmann Benjamin\01_PhD\`
The `I:` drive maps to this location.
