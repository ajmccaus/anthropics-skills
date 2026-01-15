# xlsx and docx Skills: Windows 11 Compatibility & Efficiency Assessment

## Executive Summary

The xlsx and docx skills have significant gaps in Windows 11 support and contain efficiency issues that compound on Windows due to slower file I/O. This assessment identifies specific problems and provides improvement options.

---

## Part 1: Windows 11 Compatibility Issues

### xlsx Skill

#### 1.1 LibreOffice Path Detection (Critical)

**File:** `skills/xlsx/recalc.py:18-21`

**Current Code:**
```python
if platform.system() == 'Darwin':
    macro_dir = os.path.expanduser('~/Library/Application Support/LibreOffice/4/user/basic/Standard')
else:
    macro_dir = os.path.expanduser('~/.config/libreoffice/4/user/basic/Standard')
```

**Problem:** The `else` branch assumes Linux. Windows uses `%APPDATA%\LibreOffice\4\user\basic\Standard`.

**Options:**
| Option | Description | Complexity |
|--------|-------------|------------|
| A. Add Windows path | Add `elif platform.system() == 'Windows'` with `APPDATA` path | Low |
| B. Registry lookup | Query Windows registry for LibreOffice install location | Medium |
| C. Auto-detect | Search common installation paths on all platforms | Medium |

---

#### 1.2 No Timeout Protection on Windows (Major)

**File:** `skills/xlsx/recalc.py:79-90`

**Current Code:**
```python
if platform.system() != 'Windows':
    timeout_cmd = 'timeout' if platform.system() == 'Linux' else None
    # ... gtimeout for macOS ...
    if timeout_cmd:
        cmd = [timeout_cmd, str(timeout)] + cmd
```

**Problem:** Windows is explicitly excluded from timeout handling. A hung LibreOffice process will block indefinitely.

**Options:**
| Option | Description | Complexity |
|--------|-------------|------------|
| A. subprocess timeout | Use Python's `subprocess.run(..., timeout=X)` for all platforms | Low |
| B. threading watchdog | Use `threading.Timer` to kill process after timeout | Medium |
| C. PowerShell wrapper | Use PowerShell `Start-Process -Wait -Timeout` on Windows | Low |

---

#### 1.3 soffice Command Path (Major)

**File:** `skills/xlsx/recalc.py:72-76, 31`

**Problem:** Uses bare `soffice` command. On Windows:
- LibreOffice is typically at `C:\Program Files\LibreOffice\program\soffice.exe`
- Not automatically in PATH
- May need `.exe` extension explicitly

**Options:**
| Option | Description | Complexity |
|--------|-------------|------------|
| A. PATH requirement | Document that users must add LibreOffice to PATH | Low |
| B. Auto-detect path | Search common Windows install locations | Medium |
| C. Config file | Allow users to specify LibreOffice path in config | Low |

---

#### 1.4 Documentation Linux-Only (Minor)

**File:** `skills/xlsx/SKILL.md:71`

**Current:** "Works on both Linux and macOS"

**Problem:** Windows not mentioned, making users uncertain if it's supported.

---

### docx Skill

#### 1.5 Validation Uses soffice (Major)

**File:** `skills/docx/ooxml/scripts/pack.py:103-116`

**Problem:** Same `soffice` path issues as xlsx. Validation will fail on Windows if LibreOffice isn't in PATH.

---

#### 1.6 Dependencies Use apt-get (Minor)

**File:** `skills/docx/SKILL.md:190-197`

**Current:**
```bash
sudo apt-get install pandoc
sudo apt-get install libreoffice
sudo apt-get install poppler-utils
```

**Problem:** Linux-only instructions. Windows users need:
- `winget install pandoc` or Chocolatey
- LibreOffice installer from website
- Poppler Windows binaries (non-trivial to install)

**Options:**
| Option | Description | Complexity |
|--------|-------------|------------|
| A. Multi-platform docs | Add Windows installation instructions | Low |
| B. Auto-installer script | Create setup script for each platform | Medium |
| C. Docker container | Provide containerized environment | High |

---

#### 1.7 Shebang Lines (Minor)

**Files:** All Python scripts start with `#!/usr/bin/env python3`

**Problem:** Shebang is ignored on Windows. Users must invoke with `python script.py` explicitly.

**Note:** This is typically fine as long as documentation reflects it.

---

## Part 2: Efficiency Improvements

### xlsx Skill

#### 2.1 Double Workbook Loading (Critical)

**File:** `skills/xlsx/recalc.py:103, 140`

**Problem:** Loads workbook twice with different `data_only` settings.

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Single load + formula tracking | Load once with `data_only=False`, track which cells have formulas separately | 50% faster | Medium |
| B. Lazy second load | Only load second time if first pass finds no errors | Variable | Low |
| C. openpyxl optimization | Use `read_only=True` mode for the data_only pass | 30% faster | Low |

---

#### 2.2 Unbounded Cell Iteration (Critical)

**File:** `skills/xlsx/recalc.py:112-120, 142-147`

**Problem:** Iterates ALL cells including empty ones.

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Use dimensions | Use `ws.dimensions` to get actual data range | 10-100x faster | Low |
| B. min/max bounds | Use `ws.iter_rows(min_row, max_row, min_col, max_col)` | 10-100x faster | Low |
| C. Regex error scan | Load as text and use regex for error patterns | Very fast | Medium |

---

#### 2.3 Error Pattern Matching (Moderate)

**File:** `skills/xlsx/recalc.py:115-116`

**Current:**
```python
for err in excel_errors:  # 7 iterations per cell
    if err in cell.value:
```

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Compiled regex | Use `re.compile(r'#(VALUE!|DIV/0!|REF!|...)')` | 3-5x faster | Low |
| B. Set lookup | Use `any(err in cell.value for err in error_set)` with early exit | 2x faster | Low |
| C. startswith check | Errors start with `#`, check that first | 2x faster | Low |

---

### docx Skill

#### 2.4 DOM Query Caching (Critical)

**File:** `skills/docx/scripts/document.py:75-87`

**Problem:** `_get_next_change_id()` scans entire DOM on every call.

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Instance variable cache | Cache max_id, increment on each use | 100x faster for many changes | Low |
| B. Lazy evaluation | Only scan DOM once, track additions | 100x faster | Low |

---

#### 2.5 Repeated DOM Queries in Attribute Injection (Critical)

**File:** `skills/docx/scripts/document.py:226-238`

**Problem:** Calls `getElementsByTagName()` multiple times for same tags.

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Single traversal | Walk DOM once, dispatch by tag name | 5-10x faster | Medium |
| B. Cache queries | Store results of `getElementsByTagName()` calls | 3-5x faster | Low |
| C. XPath | Use XPath for more efficient queries | Variable | Medium |

---

#### 2.6 Double Directory Scan (Moderate)

**File:** `skills/docx/ooxml/scripts/unpack.py:20`

**Current:**
```python
xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
```

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Single glob | Use `rglob("*")` and filter by extension | 2x faster | Low |
| B. os.walk | Single pass with `os.walk()` | 2x faster | Low |

---

#### 2.7 Inefficient Tag Filtering (Moderate)

**File:** `skills/docx/ooxml/scripts/pack.py:139`

**Current:**
```python
for element in dom.getElementsByTagName("*"):
    if element.tagName.endswith(":t"):
        continue
```

**Options:**
| Option | Description | Impact | Complexity |
|--------|-------------|--------|------------|
| A. Query specific tags | Only query for tags we need to process | 3-5x faster | Low |
| B. Set-based filtering | Build set of tags to skip, check membership | 2x faster | Low |

---

## Part 3: Recommended Implementation Priority

### Phase 1: Windows 11 Critical Fixes
1. Add Windows LibreOffice path detection (1.1)
2. Implement cross-platform timeout using subprocess (1.2)
3. Add soffice path detection/configuration (1.3)

### Phase 2: Efficiency Quick Wins
1. Cache `_get_next_change_id()` result (2.4)
2. Use worksheet dimensions for cell iteration (2.2)
3. Compile regex for error patterns (2.3)
4. Single directory scan in unpack.py (2.6)

### Phase 3: Documentation & UX
1. Add Windows installation instructions
2. Update SKILL.md to mention Windows support
3. Add troubleshooting section for common Windows issues

### Phase 4: Advanced Optimizations
1. Single DOM traversal for attribute injection (2.5)
2. Consolidate workbook loading (2.1)
3. Optimize tag filtering in pack.py (2.7)

---

## Windows 11 Specific Considerations

### File System Performance
Windows NTFS is slower than Linux ext4 for many small file operations (common in OOXML manipulation). Efficiency improvements have amplified benefit on Windows.

### Path Length Limits
Windows has a 260-character path limit by default. Deep OOXML structures + long filenames can hit this limit. Consider:
- Using `\\?\` prefix for long paths
- Keeping temp directories shallow

### Antivirus Interference
Windows Defender may scan every file operation, slowing document processing significantly. Consider:
- Documenting antivirus exclusion recommendations
- Batching file operations to reduce scan triggers

### LibreOffice Installation Locations
Common Windows paths to check:
```
C:\Program Files\LibreOffice\program\soffice.exe
C:\Program Files (x86)\LibreOffice\program\soffice.exe
%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe
```

### Alternative to LibreOffice on Windows
For xlsx recalculation, consider:
- **Option:** Use Excel COM automation via `win32com` when available
- **Benefit:** Native Windows integration, faster, more reliable
- **Drawback:** Requires Excel installation

---

## Summary

| Category | Issues Found | High Priority | Medium Priority |
|----------|-------------|---------------|-----------------|
| Windows Compatibility | 7 | 3 | 4 |
| Efficiency | 7 | 4 | 3 |
| **Total** | **14** | **7** | **7** |

The xlsx and docx skills require moderate work to fully support Windows 11. The efficiency improvements will benefit all platforms but are especially impactful on Windows due to slower file I/O characteristics.
