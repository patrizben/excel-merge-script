import os
import re
import subprocess
import tempfile

# ── CONFIGURATION ──────────────────────────────────────────────────────────────
INPUT_FOLDER = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\v2"
OUTPUT_FILE  = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\PEUJun2026_storylinechartsv2.xlsx"
# ───────────────────────────────────────────────────────────────────────────────

def natural_sort_key(filename):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', filename)]

# Collect files in natural sort order
files = sorted(
    [f for f in os.listdir(INPUT_FOLDER) if f.endswith((".xlsx", ".xlsm"))],
    key=natural_sort_key
)

# Build a PowerShell-safe list of files
files_ps = ",\n    ".join(f'"{f}"' for f in files)

ps_script = f"""
$inputFolder = "{INPUT_FOLDER}"
$outputFile  = "{OUTPUT_FILE}"
$files       = @(
    {files_ps}
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible       = $false
$excel.DisplayAlerts = $false

function Unique-TabName($existingTitles, $rawName) {{
    $tabName = $rawName.Substring(0, [Math]::Min(31, $rawName.Length))
    $unique  = $tabName
    $counter = 1
    while ($existingTitles -contains $unique) {{
        $suffix  = " ($counter)"
        $unique  = $tabName.Substring(0, [Math]::Min(31 - $suffix.Length, $tabName.Length)) + $suffix
        $counter++
    }}
    return $unique
}}

try {{
    $mergedWb = $excel.Workbooks.Add()
    while ($mergedWb.Sheets.Count -gt 1) {{
        $mergedWb.Sheets.Item($mergedWb.Sheets.Count).Delete()
    }}
    $mergedWb.Sheets.Item(1).Name = "__placeholder__"

    $summaryData = @()

    foreach ($filename in $files) {{
        $sourcePath = Join-Path $inputFolder $filename
        $sourceWb   = $excel.Workbooks.Open($sourcePath, $false, $true)
        $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($filename)

        for ($i = 1; $i -le $sourceWb.Sheets.Count; $i++) {{
            $sourceWs = $sourceWb.Sheets.Item($i)

            $existingTitles = @()
            for ($j = 1; $j -le $mergedWb.Sheets.Count; $j++) {{
                $existingTitles += $mergedWb.Sheets.Item($j).Name
            }}

            $rawName   = "$baseName - $($sourceWs.Name)"
            $uniqueTab = Unique-TabName $existingTitles $rawName

            $sourceWs.Copy([System.Reflection.Missing]::Value, $mergedWb.Sheets.Item($mergedWb.Sheets.Count))
            $newWs      = $mergedWb.Sheets.Item($mergedWb.Sheets.Count)
            $newWs.Name = $uniqueTab

            if ($newWs.Cells.Item(1,1).Value -ne $null -and $newWs.Cells.Item(1,1).Value -ne "") {{
                $newWs.Rows.Item(1).Insert() | Out-Null
            }}

            $a1 = $newWs.Cells.Item(1,1)
            $a1.Value = "Back to Summary"
            $newWs.Hyperlinks.Add($a1, "", "Summary!A1", "", "Back to Summary") | Out-Null
            $a1.Font.Color     = 0xFF0000  # Blue in BGR
            $a1.Font.Underline = $true

            $summaryData += ,@($uniqueTab, $filename)
            Write-Host "  Added: '$uniqueTab' from '$filename'"
        }}

        $sourceWb.Close($false)
    }}

    # Summary sheet
    $mergedWb.Sheets.Add($mergedWb.Sheets.Item(1)) | Out-Null
    $summaryWs      = $mergedWb.Sheets.Item(1)
    $summaryWs.Name = "Summary"

    $summaryWs.Cells.Item(1,1).Value     = "Summary of Merged Sheets"
    $summaryWs.Cells.Item(1,1).Font.Bold = $true
    $summaryWs.Cells.Item(1,1).Font.Size = 14
    $summaryWs.Cells.Item(3,1).Value     = "Sheet Tab"
    $summaryWs.Cells.Item(3,1).Font.Bold = $true
    $summaryWs.Cells.Item(3,2).Value     = "Source File"
    $summaryWs.Cells.Item(3,2).Font.Bold = $true

    $row = 4
    foreach ($entry in $summaryData) {{
        $tabName    = $entry[0]
        $sourceFile = $entry[1]
        $cell       = $summaryWs.Cells.Item($row, 1)
        $cell.Value = $tabName
        $summaryWs.Hyperlinks.Add($cell, "", "'$tabName'!A1", "", $tabName) | Out-Null
        $cell.Font.Color     = 0xFF0000
        $cell.Font.Underline = $true
        $summaryWs.Cells.Item($row, 2).Value = $sourceFile
        $row++
    }}

    $summaryWs.Columns.Item("A").ColumnWidth = 40
    $summaryWs.Columns.Item("B").ColumnWidth = 35

    $mergedWb.Sheets.Item("__placeholder__").Delete()

    if (Test-Path $outputFile) {{ Remove-Item $outputFile }}
    $mergedWb.SaveAs($outputFile, 51)
    $mergedWb.Close($false)
    Write-Host "`nDone! Saved to: $outputFile"

}} finally {{
    $excel.DisplayAlerts = $true
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

# Write the PowerShell script to a temp file and execute it
with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
    ps_path = f.name
    f.write(ps_script)

try:
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path],
        capture_output=True,
        text=True
    )
    print(result.stdout)
    if result.stderr:
        print("ERRORS:\n", result.stderr)
finally:
    os.remove(ps_path)
