## Change log
----------------------

Version 5.2-SNAPSHOT
-------------

## [Unreleased]

### 🚀 Gradle Improvements
- Refactored repository configuration to use `setName()` and `setUrl()`
- Moved Java compilation settings to `gradle/java-compile.gradle`
- Introduced external exclusion lists:
  - `gradle/list/exclude-license-files.list` (for licensing exclusions)
  - `gradle/list/exclude-gradle-files.list` (for Gradle file exclusions)
- Ensured `test` task runs **after** `jar`
- Improved Gradle publishing configuration with better credential handling

### 📄 Markdown to DOCX Enhancements
- **Removed** outdated Markdown-to-DOCX converters
- **Added** `Md2DocxConverter`:
  - Uses **CommonMark** for Markdown parsing
  - Generates formatted **DOCX files** with proper headings, lists, and styles
- **Introduced** `MarkdownToKdpDocxConverter`:
  - Optimized DOCX output for **Kindle Direct Publishing (KDP)**
- **Added** `MarkdownToHtmlConverter`:
  - Uses **Flexmark & JSoup** for **Markdown to HTML** conversion
- **Unit Tests**:
  - Added tests for **Markdown to DOCX**
  - Added tests for **Markdown to HTML**

### 🔄 Dependency Updates
- **Added:** `jsoup 1.18.3` for HTML processing

### Changed
- Upgraded Gradle wrapper from `8.10.2` to `8.13-RC-1`.
- Updated dependencies:
    - `assertj-core` to `3.27.3`
    - `commons-text` to `1.13.0`
    - `file-worker` to `19.0`
    - `mockito-core` to `5.15.2`
    - `poi-ooxml` and `poi` to `5.4.0`
    - `junit-jupiter` to `5.12.0-RC2`
- Introduced dependency bundles in `dependencies.gradle` to improve structure.

Version 5.1
-------------

ADDED:

- export statement for export the main package
- new method that export the given content to an Excel file
- new factory method for create a SXSSFWorkbook object

Version 5.0
-------------

CHANGED:

- repository owner
- update to jdk version 21
- update of gradle to new version 8.10.2
- update of dependency lombok to new version 1.18.34
- update of dependency poi to new version 5.3.0
