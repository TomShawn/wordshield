# WordShield

WordShield is an offline macOS desktop application for redacting and restoring confidential text in Microsoft Word documents.

The user marks sensitive text manually in Word with the yellow **Text Highlight Color**. WordShield detects those highlighted regions, replaces them with stable placeholders such as `【SECRET_001】`, and stores the original text in a password-encrypted local mapping file. After the redacted document has been edited, WordShield can restore the original text and produce an anomaly report.

WordShield does not attempt to discover sensitive entities automatically. The person reviewing the document remains responsible for marking and checking all confidential content.

Current version: **0.2.0**

## Key properties

- Fully local document processing with no cloud API, telemetry, or document-content logging
- Yellow-highlight detection across split Word runs
- Support for normal body text, nested tables, headers, footers, footnotes, and endnotes
- Preview and manual confirmation before export
- Reuse of placeholders for identical text, with manual separation for same-name entities
- Encrypted mapping files using Scrypt and AES-256-GCM
- Placeholder restoration with missing, unknown, modified, split, and count-change checks
- DOCX package inspection, metadata removal, and residual-text scanning
- Chinese/English Tkinter interface with instant language switching
- External development environment and build output, keeping the Git checkout clean

## Quick workflow

1. Open a copy of the source document in Word.
2. Mark every sensitive item with yellow **Text Highlight Color**.
3. Save the document as a normal `.docx` file.
4. Open WordShield and select **脱敏 Word**.
5. Review every detected item, entity group, location, and replace/retain decision.
6. Export the redacted DOCX, encrypted `.wsmap`, and JSON inspection report.
7. Edit the redacted document locally while preserving every placeholder exactly.
8. Open **回填 Word**, select the edited DOCX and its matching `.wsmap`, and enter the password.
9. Review the restoration report and inspect the final document in Word.

## Run the packaged application

The packaged application is a complete `WordShield.app`. It includes Python, Tkinter, and the required libraries, so end users do not need to install Python.

Double-click `WordShield.app` in Finder. You may copy the entire application bundle to `/Applications`, but keep the bundle intact.

Use the **语言 / Language** selector in the upper-right corner to switch between Chinese and English. The switch immediately updates tabs, buttons, preview headings, structural findings, dialogs, status messages, and restoration results. It applies only to the current session and creates no preference or cache file in the project or application directory.

Preview selection, entity groups, retained-text decisions, and confirmations persist when switching languages. Native macOS dialog buttons and system menus follow the operating system's language settings.

The current build is for Apple Silicon (`arm64`) and uses an ad-hoc local signature. It has not been notarized with an Apple Developer ID. A separate build is required for Intel Macs. The current acceptance environment is macOS 27.0, so other macOS versions should first be tested with fictional material.

If macOS blocks an unnotarized copy, verify its source with the maintainer or your organization administrator before using it.

## Prepare an offline workspace

Choose a local folder that is not synchronized by iCloud, OneDrive, Dropbox, or another sync client. Whether Desktop and Documents are synchronized depends on the Mac's configuration.

For a physically disconnected workflow, disable Wi-Fi, Ethernet, and other network connections before processing files. WordShield itself does not perform network requests, but Microsoft Word, backup software, sync clients, and the operating system have their own behavior and configuration.

If the entire workflow must remain offline, edit the redacted document with local Word or an approved offline AI environment. The editing stage is outside WordShield.

## Mark sensitive text in Word

1. Open a copy of the source document.
2. Select a sensitive name, location, organization, amount, identifier, or other confidential text.
3. In Word's **Home** tab, choose yellow **Text Highlight Color**.
4. Continue through body text, tables, headers, and footers.
5. Save the document as `.docx` before importing it into WordShield.

Use text highlighting rather than font color, paragraph shading, or table-cell fill. Those are different Word formatting features.

Highlight as precisely as possible. Punctuation and separator spaces should normally remain outside the highlighted region.

Word may divide one visible name into multiple internal runs. WordShield merges contiguous yellow text within the same paragraph. Paragraph boundaries, tabs, line breaks, and embedded objects remain boundaries.

## Redact a document

### Select and inspect the DOCX

Open **脱敏 Word** and click **选择 DOCX 并识别**. After local analysis, the preview shows the detected text and any unmarked occurrences of the same text.

The preview columns mean:

| Column | Meaning |
| --- | --- |
| 敏感文字 | Detected text. Display quotation marks show the exact boundary; spaces are part of the value. |
| 实体分组 | Items in the same group use the same placeholder. |
| 来源 / 处理 | Shows yellow-highlight or unmarked-match origin and whether the item will be replaced or retained. |
| 位置 | DOCX part, paragraph number, and character position for manual verification. |

An **未标记同文** item means highlighted text also appears in an unhighlighted location. WordShield replaces those matching locations by default. Review each one because identical text may represent a different real-world entity.

### Adjust groups and highlighted regions

| Button | Purpose |
| --- | --- |
| 修改所选分组 | Edit the group of the selected item. Double-clicking a row has the same effect. |
| 同名异实体单独分组 | Give one occurrence of an identical name a separate placeholder. |
| 拆分连续高亮 | Divide one contiguous yellow region into multiple items. |
| 保留／恢复所选原文 | Retain punctuation, an accidental highlight, or explicitly approved public text. Click again to restore replacement. |

Different text values cannot share one group.

Split positions mean “split after this character.” Multiple positions use comma-separated integers. Spaces and punctuation count as characters.

Examples:

- For `林知远周映禾`, enter `3` to split after the first three characters.
- For `林知远、周映禾`, enter `3,4`, then select the punctuation row and click **保留／恢复所选原文**.

Retained items appear in gray. Clearing **同时替换预览中的未标记同文位置** retains all unmarked matches. Retained original text will remain in the exported document, so WordShield requires a separate confirmation before export.

### Review structural findings

The inspection area reports DOCX structures that need attention.

- **需审查** means the user must inspect the corresponding Word content before accepting the remaining risk.
- **必须修复** means the source must be corrected in Word and imported again.
- Yellow-highlighted text in an unsupported area is blocking and cannot be bypassed by accepting risk.
- Any unapproved residual occurrence of a known sensitive value blocks export.

Text boxes, comments, deleted revisions, field instructions, images, formulas, embedded objects, hidden text, and external links may require manual review or cleanup. Yellow text in unsupported locations must first be moved into supported document text or removed in Word.

### Export

1. Select **我已核对全部敏感文字、位置和分组，确认导出** after reviewing the preview.
2. If structural warnings exist, select the risk confirmation only after completing the manual review.
3. Click **设置密码并导出脱敏文件与加密映射表**.
4. Confirm every retained-original location if prompted.
5. Enter and confirm a password containing at least 12 characters.
6. Choose a new filename in a local folder outside the source checkout and application installation directory.

WordShield creates three files next to each other:

| File | Purpose |
| --- | --- |
| `name.docx` | Redacted Word document containing placeholders |
| `name.wsmap` | Password-encrypted local mapping required for restoration |
| `name.检查报告.json` or `name.inspection-report.json` | Replacement counts, retained locations, and structural findings. The filename and human-readable report values follow the selected language. |

Keep the mapping and password separately. A lost password cannot be recovered by WordShield. The report omits entity text, group text, and source paths.

WordShield will not overwrite existing outputs. Use a new output name if any associated DOCX, mapping, or report already exists.

Open the redacted DOCX in Word and inspect body text, tables, headers, footers, line wrapping, and pagination before sharing it.

## Edit the redacted document

Keep placeholders such as `【SECRET_001】` exactly intact. Do not change their brackets, case, underscore, or number.

A useful instruction for an editor or offline AI system is:

> Preserve every placeholder exactly, including brackets, capitalization, underscores, and numbers. Reuse the corresponding placeholder when referring to the same entity. If content containing a placeholder is removed, list the removed placeholders for manual review.

Save the edited result as a local DOCX before restoration.

## Restore confidential text

1. Open **回填 Word**.
2. Select the edited DOCX under **AI 定稿 DOCX**. The same workflow applies to manually edited documents.
3. Select the corresponding `.wsmap` under **加密映射表**.
4. Click **检查定稿并回填**.
5. Review structural warnings and inspect the document in Word when required.
6. Enter the mapping password.
7. Save the restored result under a new name outside the project and application directories.
8. Read the on-screen report and the generated JSON report.
9. Click **打开最近生成的回填文档** and verify the final content and formatting in Word.

The restored output contains confidential text again and must be handled according to the original material's security requirements.

### Restoration statuses

| Status or finding | Meaning |
| --- | --- |
| 通过 | Placeholder checks passed and no listed structural warning was found. Manual content and layout review is still required. |
| 回填通过，文档需审查 | Placeholder restoration passed, but structures such as images, fields, or links need manual review. |
| 需人工检查 | A placeholder anomaly requires manual investigation. |
| 未知占位符 | The final document contains a placeholder not present in this mapping. |
| 定稿中消失的占位符 | An original placeholder was deleted or modified. WordShield reports it and does not guess a restoration location. |
| 出现次数变化 | A placeholder's occurrence count differs from the redacted source. |
| 疑似修改或拆开 | A placeholder may contain changed case, spaces, spelling changes, zero-width characters, or paragraph breaks. |
| 仍含占位符或疑似标记的内部部件 | A DOCX internal part still contains a placeholder or suspicious marker. |

Correct anomalies in the edited placeholder document, save it, and run restoration again using a new output filename.

Generic placeholder numbers cannot prove that an arbitrary final document and mapping belong together. The user must verify that relationship.

## Common problems

### No yellow highlight was detected

Confirm that the text uses yellow **Text Highlight Color**, save the DOCX, and select the correct saved file.

### One preview item contains several names

They form one uninterrupted highlighted region. Mark them more precisely in Word or use **拆分连续高亮**.

### The document contains content that cannot be checked reliably

Inspect the reported images, text boxes, comments, revisions, fields, or embedded objects in Word. Blocking findings must be corrected before export.

### The source file changed after preview

Choose **选择 DOCX 并识别** again and repeat the preview review. WordShield refuses to export from stale analysis.

### The selected output folder is rejected

Choose a separate local documents folder. WordShield prevents output inside a source checkout or the installed application directory so those directories remain clean.

### The password is rejected

Confirm the mapping belongs to this material and check case, spaces, and the saved password. Authenticated encryption detects both a wrong password and mapping corruption.

### Fix formatting or pagination changes after restoration

Text length affects wrapping, table height, and pagination. Mixed formatting within one entity cannot always be restored character by character. Review and adjust the final DOCX in Word.

## Developer setup

Development requires macOS and Python 3.11 or later with Tkinter. Some Homebrew Python builds do not include `_tkinter`.

Check the base interpreter:

```sh
python3 -B -c 'import tkinter; print(tkinter.TkVersion)'
```

Create the external development environment and run the source application:

```sh
scripts/dev.sh setup
scripts/dev.sh run
```

After setup, Finder users may double-click `Launch WordShield.command`.

Available development commands:

```sh
scripts/dev.sh setup      # Create the external virtual environment and install dependencies
scripts/dev.sh run        # Run the application from this checkout
scripts/dev.sh test       # Run the automated test suite
scripts/dev.sh fixtures   # Generate fictional DOCX fixtures outside the checkout
scripts/dev.sh build      # Build and ad-hoc sign WordShield.app outside the checkout
scripts/dev.sh paths      # Print the checkout and external development paths
```

`WORDSHIELD_PYTHON` may specify a base Python with Tkinter. `WORDSHIELD_DEV_HOME` may specify another development directory, but it must remain outside the source checkout.

The default external location is:

```text
~/Library/Caches/WordShield-Development/<checkout-id>/
  venv/                  Python development environment
  fixtures/              Generated fictional DOCX material
  build/                 PyInstaller intermediate files
  dist/                  Built WordShield.app
  pyinstaller-cache/     PyInstaller cache
  legacy/                Historical artifacts moved from the original workspace
```

The bundled entry points disable Python bytecode output. Pytest cache generation is disabled, and PyInstaller output paths are explicitly external. When using an IDE or different commands, configure its environments, caches, coverage data, and build output outside the checkout as well.

## Repository layout

```text
wordshield/               DOCX processing, encryption integration, GUI, and path protection
scripts/                  Development, fixture generation, and macOS build scripts
tests/                    Automated tests
WordShield.spec           PyInstaller application configuration
pyproject.toml            Package metadata, dependencies, and pytest configuration
Launch WordShield.command Finder launcher for source development
TEST_RESULTS.md           Test and acceptance record
README.md                 User, developer, and release documentation
```

Generated fixtures, virtual environments, build files, application bundles, and logs live outside the checkout. `.gitignore` excludes generated files, DOCX material, encrypted mappings, reports, local configuration, and signing keys.

Before committing or pushing, inspect the working tree:

```sh
git status --short
```

Ignore rules only affect untracked files. If generated or confidential files were previously tracked, remove them from Git tracking before publishing.

## Build the application

Run:

```sh
scripts/dev.sh build
scripts/dev.sh paths
```

The build command runs the complete test suite, builds the application in the external development directory, applies an ad-hoc signature, and verifies that signature. The source checkout remains unchanged.

The project currently does not produce a ZIP archive automatically. The complete application is located under the external development directory's `dist/WordShield.app`.

## Technical design

- `python-docx` validates the principal Word package structure and creates fictional test material.
- `lxml` performs targeted DOCX XML edits while preserving paragraphs, tables, sections, and run properties.
- Contiguous yellow text is merged across runs; paragraphs, tabs, breaks, and objects remain boundaries.
- Exact repeated text is matched longest-first, with per-occurrence grouping and retain controls.
- Metadata, thumbnails, custom XML, bindings, document variables, and revision-session identifiers are removed from redacted output.
- Residual scanning separates visible/content values from recognized formatting metadata to avoid false positives on common numbers and short identifiers.
- Unsupported yellow-highlighted content is blocking. Other complex structures are reported for manual review.
- Mapping files use a random 16-byte salt, Scrypt (`N=32768`, `r=8`, `p=1`), a random 12-byte nonce, and AES-256-GCM authenticated encryption.
- Output files use exclusive creation and `0600` permissions. Recoverable multi-output write failures remove files created during that attempt.
- DOCX input is limited to 128 MB total uncompressed data, 32 MB per part, and 2,048 entries. Stored and Deflate compression are supported and actual read lengths are checked in chunks.

## Clean-checkout and confidentiality policy

The application refuses to save processed DOCX files, mappings, or reports inside a source checkout or application installation directory. Business document processing remains in memory until explicit output.

The project's supported commands keep their environments and generated files outside the repository. Automated tests cover output protection, symlink bypass attempts, external fixtures, Git ignore rules, build isolation, and source/application startup without directory changes.

Finder, IDEs, Git clients, antivirus tools, Microsoft Word, backup tools, and sync clients may create their own metadata or caches. Configure those tools separately when a strict no-write checkout is required.

## Test and verify

The current suite contains 85 automated tests covering highlight extraction, cross-run replacement, tables, headers, footers, footnotes, style inheritance, residual scanning, encryption, malformed placeholders, output rollback, unsupported structures, bilingual UI behavior, clean-directory behavior, and GUI workflows.

The latest verified build and startup check confirmed:

- 85 tests passed
- The supported test workflow did not add or modify repository files
- No source file was added, removed, or modified by those supported workflows
- No file inside the packaged application changed during standalone startup

The English interface has also been checked after a real mouse selection on an unlocked Mac. A restoration-page layout issue found during that review was fixed so the button for opening the restored document remains visible. Layout tests now require every control on the active page to be displayed at both supported window sizes.

Fictional DOCX fixtures have also been rendered locally for layout inspection. Microsoft Word acceptance testing, Apple Developer ID signing, notarization, and broader macOS-version validation remain outstanding.

## Known limitations

- WordShield only processes manual highlights and exact text matches. It cannot discover all unmarked sensitive information.
- Text inside images and compressed embedded objects cannot be inspected reliably.
- Confirmed complex structures and retained original text still require human review.
- Sensitive text exists temporarily in process memory, the input DOCX, and the restored output. macOS swap, backups, Word caches, and synchronization are outside WordShield's control.
- Mixed character formatting within one entity cannot always be reproduced exactly after restoration.
- Placeholder length changes can affect wrapping, pagination, and table row height.
- If an editor swaps two valid placeholder numbers while preserving all counts, WordShield cannot determine the semantic error automatically.
- Suspicious placeholder detection covers common changes but cannot enumerate every possible rewrite.
- Structural parsing success does not guarantee that every version of Microsoft Word will open the file without repair.
- Strict OOXML, macros, digital signatures, active binary content, malformed XML/ZIP packages, and oversized packages are rejected.
- A power loss or process crash cannot make a multi-file DOCX/mapping/report export transaction fully atomic; verify that all expected files exist after interruption.

## Security reporting and contributions

Use fictional documents for development and bug reports. Do not attach confidential DOCX files, mappings, passwords, or restored outputs to GitHub issues.

Before submitting a change:

```sh
scripts/dev.sh test
git status --short
```

Keep new caches, test artifacts, fixture output, and document samples outside the repository. Add regression tests for document-processing and security-sensitive changes.
