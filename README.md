# WordShield

WordShield is an offline macOS desktop application for redacting and restoring confidential text in Microsoft Word documents.

The user marks sensitive text manually in Word with the yellow **Text Highlight Color**. WordShield detects those highlighted regions, replaces them with stable placeholders such as `【SECRET_001】`, and stores the original text in a password-encrypted local mapping file. After the redacted document has been edited, WordShield can restore the original text and produce an anomaly report.

WordShield does not attempt to discover sensitive entities automatically. The person reviewing the document remains responsible for marking and checking all confidential content.

Current version: **0.2.3**

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

The application starts in English. Use the **语言 / Language** selector in the upper-right corner to switch between English and Chinese. The switch immediately updates tabs, buttons, preview headings, structural findings, dialogs, status messages, and restoration results. It applies only to the current session and creates no preference or cache file in the project or application directory.

Preview selection, entity groups, retained-text decisions, and confirmations persist when switching languages. Native macOS dialog buttons and system menus follow the operating system's language settings.

Release archives are named `WordShield-<version>-macOS-<architecture>.zip`. A `universal2` archive runs natively on both Apple Silicon and Intel Macs. The application targets macOS 11 or later. Version 0.2.2 passed acceptance on macOS 27.0 (Apple Silicon) and macOS 15.0.1 (Intel). Test other macOS versions with fictional material first.

The application uses an ad-hoc signature only. It is not signed with an Apple Developer ID and is not notarized, so macOS blocks a downloaded copy on first launch:

1. Verify the archive against the published SHA-256 checksum, for example `shasum -a 256 -c WordShield-0.2.3-macOS-universal2.zip.sha256`.
2. Extract the archive and move `WordShield.app` to `/Applications`.
3. Double-click the application. When macOS reports that it cannot verify the developer, close the message.
4. Open **System Settings › Privacy & Security**, find the message about WordShield, and click **Open Anyway**.

Only do this for an archive whose checksum matches the one published by the maintainer or your organization administrator.

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

Highlight as precisely as possible. Punctuation should normally remain outside the highlighted region. Leading and trailing whitespace inside a highlight is ignored automatically, so `John Smith ` and `John Smith` are treated as the same entity and the surrounding spaces stay in the document.

Word may divide one visible name into multiple internal runs. WordShield merges contiguous yellow text within the same paragraph. Paragraph boundaries, tabs (including positional tabs in headers and footers), line breaks, and embedded objects remain boundaries.

## Redact a document

### Select and inspect the DOCX

Open **脱敏 Word** and click **选择 DOCX 并识别**. After local analysis, the preview shows the detected text and any unmarked occurrences of the same text.

The preview columns mean:

| Column | Meaning |
| --- | --- |
| 敏感文字 | Detected text. Display quotation marks show the exact boundary. Leading and trailing spaces are excluded; spaces inside the text are part of the value. |
| 实体分组 | Items in the same group use the same placeholder. |
| 来源 / 处理 | Shows yellow-highlight or unmarked-match origin and whether the item will be replaced or retained. |
| 位置 | DOCX part, paragraph number, and character position for manual verification. |

An **未标记同文** item means highlighted text also appears in an unhighlighted location. WordShield replaces those matching locations by default. Review each one because identical text may represent a different real-world entity.

### Adjust groups and highlighted regions

| Button | Purpose |
| --- | --- |
| Assign item to group / 调整所选条目分组 | Assign the selected occurrence to a group. Double-clicking a row has the same effect. |
| Rename group / 重命名整个分组 | Rename all members of the selected group, including unmarked matches and retained items. |
| Separate entity / 同名异实体单独分组 | Give one occurrence of an identical name a separate placeholder. |
| Split highlight / 拆分连续高亮 | Divide one contiguous yellow region into multiple items. |
| Retain / replace text / 保留／恢复所选原文 | Retain punctuation, an accidental highlight, or explicitly approved public text. Click again to restore replacement. |

Different text values cannot share one group.

Use **Rename group** when changing an entity's group label. Other groups representing same-name entities keep their own assignments. A rename requires a non-empty name that is available, so an existing group's name is rejected. Use **Assign item to group** to change the membership of one occurrence.

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
6. Choose a new filename in a local folder outside the source checkout and the `.app` bundle.

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
7. Save the restored result under a new name outside the source checkout and the `.app` bundle.
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

Choose a local documents folder. Source checkouts and `.app` bundles are protected. Document folders beside an application in Desktop, Documents, or Downloads can receive output.

### The password is rejected

Confirm the mapping belongs to this material and check case, spaces, and the saved password. Authenticated encryption detects both a wrong password and mapping corruption.

### Fix formatting or pagination changes after restoration

Text length affects wrapping, table height, and pagination. Restored text takes the formatting of the part of the placeholder that covers the most characters. Word stores the brackets of `【SECRET_001】` separately, so bolding only `SECRET_001` still makes the restored text bold. Mixed formatting within one placeholder, such as italicizing half of it, is not restored character by character. Review and adjust the final DOCX in Word.

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
scripts/dev.sh word-acceptance  # Check fictional fixtures with the local Microsoft Word
scripts/dev.sh build      # Build and ad-hoc sign WordShield.app outside the checkout
scripts/dev.sh paths      # Print the checkout and external development paths
```

`WORDSHIELD_PYTHON` may specify a base Python with Tkinter. `WORDSHIELD_DEV_HOME` may specify another development directory, but it must remain outside the source checkout.

The default external location is:

```text
~/Library/Caches/WordShield-Development/<checkout-id>/
  venv/                  Python development environment
  venv-universal2/       Environment for universal2 builds (optional)
  fixtures/              Generated fictional DOCX material
  build/                 PyInstaller intermediate files
  dist/                  Built WordShield.app, release ZIP, and SHA-256 checksum
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

The build command first checks that installed native extensions do not link external dynamic libraries such as Homebrew OpenSSL. It then runs the complete test suite, builds the application in the external development directory, applies an ad-hoc signature, verifies that signature, and runs the packaged executable in a self-test mode (`WORDSHIELD_SELF_TEST=1`). The self-test imports every dependency and redacts and restores a small fictional document in memory, without opening the interface or writing files. The source checkout remains unchanged.

When the Git working tree is clean, the build also creates `dist/WordShield-<version>-macOS-<architecture>.zip` with `ditto`, which preserves the bundle's symbolic links and signature, and a matching `.sha256` file. It prints the commit the archive was built from. With uncommitted changes, the application is still built, but no release archive is created. Tag the commit before publishing the archive.

The version number is maintained only in `wordshield/__init__.py` and `pyproject.toml`. `WordShield.spec` reads it from `wordshield/__init__.py`.

### Build for Apple Silicon and Intel

A default build contains only the current Mac's architecture. To build a `universal2` application that runs natively on both Apple Silicon and Intel Macs, use a python.org Python installer (which is universal2) and a separate environment:

```sh
WORDSHIELD_TARGET_ARCH=universal2 scripts/dev.sh setup
WORDSHIELD_TARGET_ARCH=universal2 scripts/dev.sh build
```

Setup installs the dependencies and runs `scripts/make_universal2.py`. That script replaces each native extension with its official `universal2` wheel from PyPI. When a package publishes only single-architecture wheels, such as `cffi`, it downloads the official `x86_64` and `arm64` wheels and merges them with `delocate-merge`. It does not compile from source. PyInstaller then refuses to build if any bundled binary lacks either architecture.

The packaged self-test can only run the current Mac's architecture. Before publishing a `universal2` archive built on an Intel Mac, run the application once on an Apple Silicon Mac, and vice versa.

### Check with Microsoft Word

With Microsoft Word installed, run:

```sh
scripts/dev.sh word-acceptance
```

The script generates the fictional fixtures in `word-acceptance` under the external development directory and drives Word through AppleScript. It checks that Word opens every fixture, including redacted and restored output, and that a document opened and saved again by Word still yields the same highlights and restores completely. It also edits the redacted document in Word: it inserts text next to placeholders, changes their formatting, appends a paragraph, deletes and duplicates placeholders, and edits with Track Changes on. It then checks the restoration result and report for each case. On the first run, macOS may ask whether to allow control of Word, and Word may ask for access to the folder. Grant both.

On Intel Macs, recent `cryptography` releases may not provide a prebuilt wheel. If pip compiles it from source, it links Homebrew OpenSSL, and PyInstaller replaces that library with the older, identically named OpenSSL bundled with python.org Python, so the packaged application fails at startup. `scripts/dev.sh setup` therefore prefers binary wheels, and the build stops early if a source-built extension is detected. To repair an existing environment, install an official wheel, for example:

```sh
"$(scripts/dev.sh paths | sed -n 's/^开发环境与构建目录 //p')/venv/bin/python" -m pip install --only-binary=:all: --force-reinstall "cryptography>=43,<51"
```


## Technical design

- `python-docx` validates the principal Word package structure and creates fictional test material.
- `lxml` performs targeted DOCX XML edits while preserving paragraphs, tables, sections, and run properties.
- Contiguous yellow text is merged across runs; paragraphs, tabs, breaks, and objects remain boundaries.
- Exact repeated text is matched longest-first, with per-occurrence grouping and retain controls.
- Metadata, thumbnails, custom XML, bindings, document variables, and revision-session identifiers are removed from redacted output.
- Relationship cleanup resolves internal targets relative to their source parts and removes references to deleted metadata, including custom relationship types.
- Residual scanning separates visible/content values from recognized formatting metadata, including picture size, position, and identifier values, to avoid false positives on common numbers and short identifiers. Picture alt text and names are still checked.
- Binary parts such as images are scanned only for encoded matches of at least four bytes, because shorter byte sequences occur by chance in image data. Very short terms, such as a two-digit number or a single CJK character, are therefore not detected inside binary content, which already requires manual review.
- Unsupported yellow-highlighted content is blocking. Other complex structures are reported for manual review.
- Mapping files use a random 16-byte salt, Scrypt (`N=32768`, `r=8`, `p=1`), a random 12-byte nonce, and AES-256-GCM authenticated encryption.
- Output files use exclusive creation and `0600` permissions. Recoverable multi-output write failures remove files created during that attempt.
- DOCX input is limited to 128 MB total uncompressed data, 32 MB per part, and 2,048 entries. Stored and Deflate compression are supported and actual read lengths are checked in chunks.

## Clean-checkout and confidentiality policy

The application protects source checkouts and `.app` bundles from processed DOCX files, mappings, and reports. The containing Desktop, Documents, or Downloads folder can be used for document output. Business document processing remains in memory until explicit output.

The project's supported commands keep their environments and generated files outside the repository. Automated tests cover output protection, symlink bypass attempts, external fixtures, Git ignore rules, build isolation, and source/application startup without directory changes.

Finder, IDEs, Git clients, antivirus tools, Microsoft Word, backup tools, and sync clients may create their own metadata or caches. Configure those tools separately when a strict no-write checkout is required.

## Test and verify

The current suite contains 161 automated tests covering highlight extraction, cross-run replacement, tables, headers, footers, footnotes, style inheritance, modern Word identifiers and formatting, metadata relationships, group renaming, residual scanning, encryption, malformed placeholders, output rollback, unsupported structures, bilingual UI behavior, clean-directory behavior, and GUI workflows.

The latest verified build and startup check confirmed:

- 161 tests passed
- The supported test workflow did not add or modify repository files
- No source file was added, removed, or modified by those supported workflows
- No file inside the packaged application changed during standalone startup

The English interface has also been checked after a real mouse selection on an unlocked Mac. A restoration-page layout issue found during that review was fixed so the button for opening the restored document remains visible. Layout tests now require every control on the active page to be displayed at both supported window sizes. Version 0.2.3 clears the restoration report when either input changes or a new inspection starts. Failed or cancelled attempts leave the report pending. Results from a background task are displayed only when its inputs still match the current fields. Regression checks cover file pickers, typed paths, same-input retries, failures, cancellation, late callbacks and language switching.

Fictional DOCX fixtures have also been rendered locally for layout inspection. Microsoft Word 16.102.3 on macOS 15.0.1 opened all 12 fictional fixtures, including redacted and restored output. After Word saved them again, highlights were still detected identically and restoration completed with matching text. Word only adds empty paragraphs after tables and footnote/endnote separators when saving, which does not affect detection or restoration. After real edits in Word, restoration kept the inserted text, bold, font, and color, reported deleted and duplicated placeholders, and required risk confirmation for tracked changes. Apple Developer ID signing, notarization, and broader macOS-version validation remain outstanding.

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
