# DEMO_DUMMY

## Shivansh / ONGC BRRP project archive

This repository contains the available project documents, generated management packages, analysis records and code from the BRRP work in Codex.

**Start with P14, the latest consolidated management-development issue:**

| File | Purpose |
| --- | --- |
| [Execution manual (PDF)](BRRP/deliverables/BRRP_P14_Project_Execution_Management_Manual.pdf) | Readable A–Z execution and management procedures |
| [Execution manual (Word)](BRRP/deliverables/BRRP_P14_Project_Execution_Management_Manual.docx) | Editable manual |
| [Integrated controls (Excel)](BRRP/deliverables/BRRP_P14_Integrated_Management_Control.xlsx) | 34 control sheets, requirements, engineering, procurement, readiness, cost and turnover |
| [Schedule and driving path (PDF)](BRRP/deliverables/BRRP_P14_Overall_Schedule_and_Driving_Path.pdf) | Readable schedule summary, Gantt and driving chain |
| [Integrated schedule (CSV)](BRRP/deliverables/BRRP_P14_Integrated_Schedule.csv) | 285-activity schedule data |
| [Handover and limitations](BRRP/CHAT_HANDOVER.md) | Context for continuing the work |

### Archive layout

- `BRRP/source-documents/`: 72 distinct project source files. Byte-identical uploads are stored once.
- [Source index](BRRP/source-index.csv) and [source manifest](BRRP/source-manifest.json): original names, duplicate aliases, checksums and archive locations; these represent 275 project upload copies.
- `BRRP/deliverables/`: all 37 available downloadable deliverables, including earlier issues.
- `BRRP/project-review/`: 400 available analysis files, scripts, inventories, extracts, intermediate results and issued packages.
- [Archive summary](BRRP/archive-summary.json) and [exclusions](BRRP/excluded-files.json): coverage and explicit omissions.

P14 is a **management-development package**, not an approved construction baseline or certified design. Earlier issues are retained as history and may contain assumptions superseded or challenged by later reviews. Source receipt does not mean every clause, value or design calculation has been independently verified.

The base schedule assumes commencement on 1 November 2026 and calculates completion on 6 April 2028. Two deterministic alternatives calculate 26 February 2028. These are conditional scenarios; actual award, resources, prices, approvals and contractual relief remain unverified. Read the manual before using the dates.

### Downloading

Open a file on GitHub and use **Download raw file** where available. To obtain the repository as a ZIP, use **Code → Download ZIP**. To retain Git history and get future updates, clone the repository with Git.

### Continuing the analysis

Historical scripts refer to `/workspace/project-review`, `/workspace/attachments` and `/workspace/downloads`. The archive preserves their code and input provenance. [Workspace restoration instructions](BRRP/tools/README.md) explain how to recreate those files without changing the originals.

The archive does not contain a verbatim chat export: no exported chat PDF was present locally. It includes a clearly labelled handover summary instead. One preliminary PDF could not be transferred because it exceeded the tool's 32 MiB limit. The unrelated personal travel ticket and Python caches were excluded.
