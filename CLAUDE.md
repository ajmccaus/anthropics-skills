# Claude Code Guidelines for This Repository

## Code Quality Standards

### Before Committing or Pushing
1. **Always validate code** - Run syntax checks and basic import tests before committing
2. **Test functions are used** - Verify new code is actually called somewhere
3. **Check for over-engineering** - If a function just wraps another function without adding value, don't create it

### Avoid Over-Engineering
- Don't create wrapper functions that add no value (e.g., wrapping `subprocess.run` with the same parameters)
- Don't write utility functions "for later" - only write what's actually needed now
- Don't add unused constants, type hints for unused code, or speculative features
- Delete code that isn't used rather than leaving it "in case it's needed"

### Code Review Checklist
Before saying work is complete:
- [ ] Syntax check passes (`python -m py_compile`)
- [ ] Imports work (`python -c "from module import func"`)
- [ ] Basic functionality tests pass
- [ ] No unused functions or dead code
- [ ] File sizes are reasonable (question large additions)

## Repository-Specific Notes

### xlsx Skill
- Uses Excel COM on Windows (preferred), LibreOffice as fallback
- `recalc.py` loads workbook twice: `data_only=True` for error detection, `data_only=False` for formula counting
- Cell iteration uses worksheet bounds to skip empty regions

### docx Skill
- Uses LibreOffice for PDF conversion and validation
- OOXML manipulation via unpacking ZIP, editing XML, repacking
