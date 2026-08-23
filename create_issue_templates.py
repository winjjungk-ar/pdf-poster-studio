from pathlib import Path

ROOT = Path.cwd()
ISSUE_DIR = ROOT / ".github" / "ISSUE_TEMPLATE"
ISSUE_DIR.mkdir(parents=True, exist_ok=True)

bug_report = """---
name: Bug report
about: Report a problem with PDF Poster Studio
title: "[Bug] "
labels: bug
assignees: ""
---

## What happened?

Describe the problem clearly.

## Steps to reproduce

1.
2.
3.

## Expected result

What did you expect to happen?

## Actual result

What actually happened?

## File type

- [ ] PDF
- [ ] PNG
- [ ] JPG / JPEG
- [ ] BMP
- [ ] WEBP
- [ ] Other:

## Workflow used

- [ ] Merge PDFs
- [ ] Extract Pages
- [ ] Make Poster
- [ ] Rotate Pages
- [ ] Selected Area PDF
- [ ] Create Poster PDF
- [ ] Create Single Page PDF
- [ ] Other:

## Output settings

Output Area:
- [ ] Whole Page
- [ ] Selected Area

Output Files:
- [ ] One PDF
- [ ] Separate PDF per selected page

Paper Size:
- [ ] A4
- [ ] US Letter

## File size / page count

Example:
- PDF size: 17 MB
- Page count: 966 pages

## Windows version

Example:
- Windows 10
- Windows 11
- Korean / Japanese / Chinese Windows, if relevant

## Error message

Paste the full error message here.

## Screenshots

Drag and drop screenshots here if possible.

## Additional notes

Anything else that may help.
"""

feature_request = """---
name: Feature request
about: Suggest an improvement or new feature
title: "[Feature] "
labels: enhancement
assignees: ""
---

## What feature do you want?

Describe the feature clearly.

## Why is it useful?

Explain the use case.

Examples:
- printing game maps
- wall posters
- engineering drawings
- merging PDFs
- extracting pages
- rotating pages
- large PDF workflows

## Current workaround

How do you handle this now?

## Suggested behavior

Describe how you think it should work.

## Screenshots or examples

Add screenshots, drawings, or example files if helpful.

## Additional notes

Anything else?
"""

config = """blank_issues_enabled: true
contact_links:
  - name: PDF Poster Studio Releases
    url: https://github.com/winjjungk-ar/pdf-poster-studio/releases
    about: Download the latest beta release.
"""

(ISSUE_DIR / "bug_report.md").write_text(bug_report, encoding="utf-8")
(ISSUE_DIR / "feature_request.md").write_text(feature_request, encoding="utf-8")
(ISSUE_DIR / "config.yml").write_text(config, encoding="utf-8")

print("[OK] Issue templates created")
print(ISSUE_DIR / "bug_report.md")
print(ISSUE_DIR / "feature_request.md")
print(ISSUE_DIR / "config.yml")
