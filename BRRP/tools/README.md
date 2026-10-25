# Restoring the recorded workspace layout

The source manifest keeps each project's upload aliases under `attachments/<upload-directory>/<filename>`. Byte-identical uploads are stored once in the repository.

To recreate the archived workspace files in a chosen empty directory, from the repository root run:

```bash
python BRRP/tools/materialize_workspace.py --root /tmp/brrp-workspace
```

The helper verifies source SHA-256 checksums, recreates the upload aliases, and restores `project-review` and `downloads`. It refuses to overwrite differing existing files. Restored files are independent copies so later edits do not change the archive. It does not run the analysis or install software.

Many historical scripts use absolute `/workspace/...` paths. To rerun those scripts unchanged, use an appropriate container with that layout, or adapt the paths deliberately. Restoring into a different directory alone does not make absolute paths portable.

`python-packages.json` records the installed versions used when this archive was prepared; it is an environment snapshot, not a guarantee that every historical script has been rerun. OCR additionally requires Tesseract and relevant system tooling.

The P14 schedule scenario generator and its verification script are in `../project-review/build_p14_schedule.py` and `../project-review/verify_p14.py`. Existing output checks are in `../project-review/outputs/BRRP_P14_Consolidated_Management/P14_quality_checks.json`. Spreadsheet formulas were not recalculated by an Excel engine in those checks.
