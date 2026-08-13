# legacy/ — the original engine-COM approach (superseded)

`Get-NeTaxCode.ps1` is the earlier design: it drove your Excel **engine
template** (`Address_Tax_Engine.xlsx`) over COM, in parallel runspaces, to get
the codes. It works, but it has a large moving-part surface (a second workbook,
runspaces, backup workbooks) and was the source of the "wrong codes / hard to
run" problems.

**Use `Get-CityCode.ps1` in the repo root instead.** It ports the same engine
formulas into pure PowerShell (`NeMatch.ps1`) — simpler, no engine workbook, and
validated to reproduce the engine's answers.

This script is kept for reference only. It depends on `NeTaxPaste.ps1` (in the
repo root) and an `Address_Tax_Engine.xlsx`; it is not wired to run from this
folder as-is.
