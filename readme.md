# 🛠️ Offset Rig Software VB

Welcome to the unified repository for the **End-of-Line (EOL) check software** running on the physical "Offset" calibration/test stations in **Unit 3**. 

This project represents the modernization, consolidation, and refactoring of legacy Visual Basic 6 (VB6) codebases migrated to the **twinBASIC** compiler ecosystem.

---

## 📌 Project Purpose & Genesis

### Origin
This project was born on **10/06/26**, forking the legacy EOL check systems.

### The Challenge
With the introduction of new products and changing quality verification standards, the EOL stations received a surge in feature requests. The legacy codebase was fragmented across separate folders for each station (Offset 1, Offset 2, Offset 3), resulting in duplicated code, hardcoded hardware variations, and severe maintainability issues.

### The Mission
The core goal of this project is two-fold:
1. **Unification:** Consolidate all physical testing stations into a **single, unified codebase** driven by configuration.
2. **Modernization:** Clean up legacy code smells, simplify data parsers, and restructure logic to prepare the software for future translation into modern language frameworks (e.g., Python or C#). Both human and AI auditors are utilized to target improvements.

---

## 🚀 Key Modernizations Achieved

*   **🎛️ Codebase Unification:** Consolidated three distinct, station-specific project directories into one repository. Station-specific hardware identifiers (VISA IDs), calibration constants, and timing delays have been abstracted into a single local config file: `offset_config.txt`.
*   **🔌 Dynamic Current Draw Overrides:** Replaced duplicated hardcoded current limits with dynamic limits loaded from `current_draw.txt` on a per-product-range basis (e.g. `RT`, `XJ`, `XV`, `FT`, `XF`).
*   **🔋 AM Product Range Exceptions:** Added exception rules (including custom voltage thresholds and limits of ±12) for the new `"AM"` product line, mapping them dynamically to the existing `"UB"` product rules.
*   **📂 Dynamic File Parsing:** Replaced the legacy fixed-length loops (e.g. `For i = 1 To 150`) with modern `Do While Not EOF(FileHandle)` structures to allow configuration lists to grow dynamically without compilation limits.

---

## 🗺️ Roadmap & Backlog (Identified by Human Auditors)

We have identified several critical areas for improvement to be resolved in upcoming sprints:

### 🏷️ A. Version Control & Release Management
*   **The Issue:** Right now, versioning is arbitrary (indicated only by `"V34"` in the main form title and `"V34.1"` in the binary's filename).
*   **The Plan:** Remove hardcoded version strings from the UI titles and filenames. Hide the active version string inside a Help/About menu. Version identifiers will correspond exactly to the **Git repository tags** under which the binaries are compiled.

### 🧹 B. Code Smells & Refactoring
1.  **DMM Switching Redundancy (`DigitalMultimeterControl.bas`):** Many switching subroutines replicate identical commands and logic solely to send command strings to the multimeter. These will be abstracted into a single parameterized router command.
2.  **Overcomplicated Barcode Parsing (`Barcode.bas`):** The parsing engine for works order sheet barcodes is excessively complex and brittle. It will be refactored into a clean, pattern-matching regex/substring scanner.
3.  **Inefficient Load Checking (`OffsetCheck.bas`):** The `CheckLoad` routine features an inefficient and convoluted conversion logic that will be rewritten for clarity and speed.

### 🗺️ C. Hardcoded Product Process Variations
*   **The Issue:** Operators cannot easily tell which testing process a particular product range will follow because routing logic is hardcoded inside branching conditional statements.
*   **The Plan:** 
    *   Clearly define and document the possible EOL testing processes.
    *   Determine the most common test path and establish it as the default system behavior.
    *   Create a clean lookup dictionary of product range overrides to route non-standard products to their respective custom tests.

---

## 🤖 AI-Identified Issues (Architectural & Code Quality Audit)

During the codebase unification process, our AI static analysis tool audited the source files and identified several critical code smells, logic errors, and architectural vulnerabilities. These are listed below, ranked by severity.

### 🔴 High Severity

#### 1. Logic Nesting Bug & Dead Code in Interlock Scanner
*   **Location:** [InputsOutputs.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/InputsOutputs.bas#L15-L57)
*   **The Issue:** The `CheckInputs()` routine runs a loop `For i = 0 To 7` to scan bits. However, the entire inner logic is wrapped inside an `If i = 0 Then` block. The `ElseIf i = 1 Then` statement on line 46 is nested *inside* the `i = 0` conditional, making it completely unreachable dead code.
*   **Impact:** Any hardware signal reading intended for inputs other than bit 0 is silently ignored. Global status flags `PartPresent`, `OringPresent`, and `RestrictorPresent` are declared but never set, leaving them as dead variables.

#### 2. Hardcoded Resistor Value Mapping Table
*   **Location:** [OffsetCheck.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/OffsetCheck.bas#L2601-L3613)
*   **The Issue:** The `CheckLoad` subroutine contains over **1,000 lines** of hardcoded conditional mapping statements (`If LoadValue = X Then LoadValue = Y`) to match resistor indices to decimal load values.
*   **Impact:** Severe violation of the Open-Closed Principle (OCP). The software requires a full recompilation whenever a new resistor value or station configuration is introduced. It also adds massive cognitive bloat to the module.

#### 3. COM Excel Instance Abuse & Process Leaking
*   **Location:** [CSVFile.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/CSVFile.bas)
*   **The Issue:** The Excel operations (`FindPODInExcelFile`, `FindResults`, `CreateExcel`, `Update25DayHoldResult`, and `UpdateExcelWithIdResults`) create a new, separate heavy Excel COM instance (`CreateObject("Excel.Application")` or `New Excel.Application`) for every single file access.
*   **Impact:** Creates severe runtime performance overhead. If a network drop or unexpected runtime error occurs, the code jumps to error handlers that call `xl.Quit` but fail to close open workbooks or release COM references, leaving zombie `EXCEL.EXE` background processes that lock files and leak system memory.

#### 4. Shared File Concurrency Vulnerabilities
*   **Location:** [CSVFile.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/CSVFile.bas)
*   **The Issue:** Rigs read and write to the same shared Excel sheets on the network (`\\USVR8\Results\Production\Offset Check Results\`) simultaneously.
*   **Impact:** There is no concurrency protection, file lock checks, or retry logic. If two rigs write data at the exact same moment, the operation will crash with a `Permission Denied` runtime error, forcing the operator to click a MsgBox and abort the run.

---

### 🟡 Medium Severity

#### 5. Lack of Option Explicit in Key Modules
*   **Location:** [Barcode.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/Barcode.bas), [CSVFile.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/CSVFile.bas), [Vision.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/Vision.bas), and [ieeevb.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/ieeevb.bas)
*   **The Issue:** These files are missing the `Option Explicit` compiler directive.
*   **Impact:** The twinBASIC compiler will not throw errors for misspelled variables, instead silently creating them as uninitialized Variant variables at runtime. This causes silent logic errors (e.g. `T4Reply` in [Vision.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/Vision.bas) is undeclared, and `IsValidAnaloguePartNumber` is used as an undeclared local variable in `SECONDMCSBARCODE`).

#### 6. Extreme Relay Switching Code Duplication
*   **Location:** [DigitalMultimeterControl.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/DigitalMultimeterControl.bas)
*   **The Issue:** Over 100 near-identical routines are written solely to configure the switch cards for different test patterns (e.g., `SwitchSTCVsMeas`, `SwitchSTCVsMeasA`, `SwitchSTCVsMeasQ`, etc.), each hardcoding identical `OpenAllSwitches` and `Sleep` sequences.
*   **Impact:** Makes the codebase extremely large and difficult to maintain. Centralizing this into a single parameterized method (e.g., `SwitchRelays(RoutePattern As String)`) would eliminate hundreds of lines of duplicate code.

#### 7. Hardcoded Local and Network Server Paths
*   **Location:** [CSVFile.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/CSVFile.bas), [LogFile.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/LogFile.bas)
*   **The Issue:** Network paths (`\\USVR8\Results\...`), local directory paths (`C:\offset setup files\`), and print paths (`M:\system\load\Vborders.txt`) are duplicated and hardcoded directly in code.
*   **Impact:** Moving the setup files, results server, or printing batch script requires editing the code in multiple places and recompiling the application.

---

### 🟢 Low Severity

#### 8. Hardcoded Calibration Constants in PSU Controls
*   **Location:** [ControlPSU.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/ControlPSU.bas#L30)
*   **The Issue:** Hardware calibration compensation values (`Supply * 0.986` and `Supply * 0.9645`) are hardcoded into the PSU voltage settings.
*   **Impact:** Since these scaling factors compensate for wiring resistance and voltage drops, they are station-dependent. They should be configured in `offset_config.txt` rather than hardcoded in source.

#### 9. Brittle Magic-Index Barcode Parsing
*   **Location:** [OffsetCheck.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/OffsetCheck.bas) and [Barcode.bas](file:///C:/Users/lewis.heslop/Downloads/Offset%20Rig%20Software%20VB/Barcode.bas)
*   **The Issue:** Barcode parameters are parsed using hardcoded substring indices (e.g. `Mid$(MainForm.SecondMCSBarcode, 17, 4)`) without any bounds checking or format validation.
*   **Impact:** Any changes to the printed barcode layout will silently parse incorrect pressure and scaling limits, risking false calibration runs without throwing format errors.

