<<<<<<< HEAD
import os
import subprocess
import tempfile

# ── CONFIGURATION ──────────────────────────────────────────────────────────────
INPUT_FOLDER = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\v2"
OUTPUT_FILE  = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\PEUJun2026_storylinechartsv2.xlsx"
# ───────────────────────────────────────────────────────────────────────────────

ps_script = f"""
$inputFolder = "{INPUT_FOLDER}"
$outputFile  = "{OUTPUT_FILE}"

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

function Sheet-HasChart($ws) {{
    foreach ($shape in $ws.Shapes) {{
        if ($shape.HasChart) {{ return $true }}
    }}
    if ($ws.ChartObjects().Count -gt 0) {{ return $true }}
    return $false
}}

function Convert-FormulasToValues($ws) {{
    $used = $ws.UsedRange
    if ($used -eq $null) {{ return }}
    $used.Copy() | Out-Null
    $used.PasteSpecial(-4163) | Out-Null
    $ws.Application.CutCopyMode = $false
}}

function Get-NaturalSortKey($str) {{
    $parts  = [regex]::Split($str, '(\\d+)')
    $result = @()
    foreach ($part in $parts) {{
        if ($part -match '^\\d+$') {{ $result += [int]$part }}
        else {{ $result += $part.ToLower() }}
    }}
    return $result
}}

try {{
    $mergedWb = $excel.Workbooks.Add()
    while ($mergedWb.Sheets.Count -gt 1) {{
        $mergedWb.Sheets.Item($mergedWb.Sheets.Count).Delete()
    }}
    $mergedWb.Sheets.Item(1).Name = "__placeholder__"

    $summaryData = @()

    $files = Get-ChildItem -Path $inputFolder -File |
             Where-Object {{ $_.Extension -match '\\.(xlsx|xlsm)$' }} |
             Sort-Object {{ (Get-NaturalSortKey $_.Name) -join '|' }}

    foreach ($file in $files) {{
        $sourcePath = $file.FullName
        $filename   = $file.Name
        $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($filename)

        Write-Host "Opening: $filename"
        $sourceWb = $excel.Workbooks.Open($sourcePath, $false, $true)

        for ($i = 1; $i -le $sourceWb.Sheets.Count; $i++) {{
            $sourceWs = $sourceWb.Sheets.Item($i)

            if (-not (Sheet-HasChart $sourceWs)) {{
                Write-Host "  Skipping '$($sourceWs.Name)' -- no charts"
                continue
            }}

            $existingTitles = @()
            for ($j = 1; $j -le $mergedWb.Sheets.Count; $j++) {{
                $existingTitles += $mergedWb.Sheets.Item($j).Name
            }}
            $rawName   = "$baseName - $($sourceWs.Name)"
            $uniqueTab = Unique-TabName $existingTitles $rawName

            $sourceWs.Copy([System.Reflection.Missing]::Value, $mergedWb.Sheets.Item($mergedWb.Sheets.Count))
            $newWs      = $mergedWb.Sheets.Item($mergedWb.Sheets.Count)
            $newWs.Name = $uniqueTab

            Convert-FormulasToValues $newWs

            if ($newWs.Cells.Item(1,1).Value -ne $null -and $newWs.Cells.Item(1,1).Value -ne "") {{
                $newWs.Rows.Item(1).Insert() | Out-Null
            }}
            $a1 = $newWs.Cells.Item(1,1)
            $a1.Value = "Back to Summary"
            $newWs.Hyperlinks.Add($a1, "", "Summary!A1", "", "Back to Summary") | Out-Null
            $a1.Font.Color     = 16711680
            $a1.Font.Underline = $true

            $summaryData += [PSCustomObject]@{{
                TabName    = $uniqueTab
                SourceFile = $filename
            }}

            Write-Host "  Added: '$uniqueTab' from '$filename'"
        }}

        $sourceWb.Close($false)
    }}

    $mergedWb.Sheets.Add($mergedWb.Sheets.Item(1)) | Out-Null
    $summaryWs      = $mergedWb.Sheets.Item(1)
    $summaryWs.Name = "Summary"

    $summaryWs.Cells.Item(1,1).Value     = "Summary of Chart Sheets"
    $summaryWs.Cells.Item(1,1).Font.Bold = $true
    $summaryWs.Cells.Item(1,1).Font.Size = 14

    $summaryWs.Cells.Item(3,1).Value     = "Sheet Tab"
    $summaryWs.Cells.Item(3,2).Value     = "Source File"
    $summaryWs.Cells.Item(3,1).Font.Bold = $true
    $summaryWs.Cells.Item(3,2).Font.Bold = $true

    $row = 4
    foreach ($entry in $summaryData) {{
        $cell       = $summaryWs.Cells.Item($row, 1)
        $cell.Value = $entry.TabName
        $summaryWs.Hyperlinks.Add($cell, "", "'$($entry.TabName)'!A1", "", $entry.TabName) | Out-Null
        $cell.Font.Color     = 16711680
        $cell.Font.Underline = $true
        $summaryWs.Cells.Item($row, 2).Value = $entry.SourceFile
        $row++
    }}

    $summaryWs.Columns.Item("A").ColumnWidth = 40
    $summaryWs.Columns.Item("B").ColumnWidth = 35

    $mergedWb.Sheets.Item("__placeholder__").Delete()

    if (Test-Path $outputFile) {{ Remove-Item $outputFile }}
    $mergedWb.SaveAs($outputFile, 51)
    $mergedWb.Close($false)

    Write-Host ""
    Write-Host "Done! Saved to: $outputFile"

}} finally {{
    $excel.DisplayAlerts = $true
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
    ps_path = f.name
    f.write(ps_script)

try:
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path],
        capture_output=True, text=True
    )
    print(result.stdout)
    if result.stderr:
        print("ERRORS:\n", result.stderr)
finally:
    os.remove(ps_path)
=======
import os
import subprocess
import tempfile

# ── CONFIGURATION ──────────────────────────────────────────────────────────────
INPUT_FOLDER = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\v2"
OUTPUT_FILE  = r"C:\Users\wb601985\OneDrive - WBG\Documents\2025-2026\3. Economic Monitoring\PEU\PEU June 2026\Storyline charts\PEUJun2026_storylinechartsv2.xlsx"
# ───────────────────────────────────────────────────────────────────────────────

ps_script = f"""
$inputFolder = "{INPUT_FOLDER}"
$outputFile  = "{OUTPUT_FILE}"

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

function Sheet-HasChart($ws) {{
    foreach ($shape in $ws.Shapes) {{
        if ($shape.HasChart) {{ return $true }}
    }}
    if ($ws.ChartObjects().Count -gt 0) {{ return $true }}
    return $false
}}

function Convert-FormulasToValues($ws) {{
    $used = $ws.UsedRange
    if ($used -eq $null) {{ return }}
    $used.Copy() | Out-Null
    $used.PasteSpecial(-4163) | Out-Null
    $ws.Application.CutCopyMode = $false
}}

function Get-NaturalSortKey($str) {{
    $parts  = [regex]::Split($str, '(\\d+)')
    $result = @()
    foreach ($part in $parts) {{
        if ($part -match '^\\d+$') {{ $result += [int]$part }}
        else {{ $result += $part.ToLower() }}
    }}
    return $result
}}

try {{
    $mergedWb = $excel.Workbooks.Add()
    while ($mergedWb.Sheets.Count -gt 1) {{
        $mergedWb.Sheets.Item($mergedWb.Sheets.Count).Delete()
    }}
    $mergedWb.Sheets.Item(1).Name = "__placeholder__"

    $summaryData = @()

    $files = Get-ChildItem -Path $inputFolder -File |
             Where-Object {{ $_.Extension -match '\\.(xlsx|xlsm)$' }} |
             Sort-Object {{ (Get-NaturalSortKey $_.Name) -join '|' }}

    foreach ($file in $files) {{
        $sourcePath = $file.FullName
        $filename   = $file.Name
        $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($filename)

        Write-Host "Opening: $filename"
        $sourceWb = $excel.Workbooks.Open($sourcePath, $false, $true)

        for ($i = 1; $i -le $sourceWb.Sheets.Count; $i++) {{
            $sourceWs = $sourceWb.Sheets.Item($i)

            if (-not (Sheet-HasChart $sourceWs)) {{
                Write-Host "  Skipping '$($sourceWs.Name)' -- no charts"
                continue
            }}

            $existingTitles = @()
            for ($j = 1; $j -le $mergedWb.Sheets.Count; $j++) {{
                $existingTitles += $mergedWb.Sheets.Item($j).Name
            }}
            $rawName   = "$baseName - $($sourceWs.Name)"
            $uniqueTab = Unique-TabName $existingTitles $rawName

            $sourceWs.Copy([System.Reflection.Missing]::Value, $mergedWb.Sheets.Item($mergedWb.Sheets.Count))
            $newWs      = $mergedWb.Sheets.Item($mergedWb.Sheets.Count)
            $newWs.Name = $uniqueTab

            Convert-FormulasToValues $newWs

            if ($newWs.Cells.Item(1,1).Value -ne $null -and $newWs.Cells.Item(1,1).Value -ne "") {{
                $newWs.Rows.Item(1).Insert() | Out-Null
            }}
            $a1 = $newWs.Cells.Item(1,1)
            $a1.Value = "Back to Summary"
            $newWs.Hyperlinks.Add($a1, "", "Summary!A1", "", "Back to Summary") | Out-Null
            $a1.Font.Color     = 16711680
            $a1.Font.Underline = $true

            $summaryData += [PSCustomObject]@{{
                TabName    = $uniqueTab
                SourceFile = $filename
            }}

            Write-Host "  Added: '$uniqueTab' from '$filename'"
        }}

        $sourceWb.Close($false)
    }}

    $mergedWb.Sheets.Add($mergedWb.Sheets.Item(1)) | Out-Null
    $summaryWs      = $mergedWb.Sheets.Item(1)
    $summaryWs.Name = "Summary"

    $summaryWs.Cells.Item(1,1).Value     = "Summary of Chart Sheets"
    $summaryWs.Cells.Item(1,1).Font.Bold = $true
    $summaryWs.Cells.Item(1,1).Font.Size = 14

    $summaryWs.Cells.Item(3,1).Value     = "Sheet Tab"
    $summaryWs.Cells.Item(3,2).Value     = "Source File"
    $summaryWs.Cells.Item(3,1).Font.Bold = $true
    $summaryWs.Cells.Item(3,2).Font.Bold = $true

    $row = 4
    foreach ($entry in $summaryData) {{
        $cell       = $summaryWs.Cells.Item($row, 1)
        $cell.Value = $entry.TabName
        $summaryWs.Hyperlinks.Add($cell, "", "'$($entry.TabName)'!A1", "", $entry.TabName) | Out-Null
        $cell.Font.Color     = 16711680
        $cell.Font.Underline = $true
        $summaryWs.Cells.Item($row, 2).Value = $entry.SourceFile
        $row++
    }}

    $summaryWs.Columns.Item("A").ColumnWidth = 40
    $summaryWs.Columns.Item("B").ColumnWidth = 35

    $mergedWb.Sheets.Item("__placeholder__").Delete()

    if (Test-Path $outputFile) {{ Remove-Item $outputFile }}
    $mergedWb.SaveAs($outputFile, 51)
    $mergedWb.Close($false)

    Write-Host ""
    Write-Host "Done! Saved to: $outputFile"

}} finally {{
    $excel.DisplayAlerts = $true
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""

with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
    ps_path = f.name
    f.write(ps_script)

try:
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path],
        capture_output=True, text=True
    )
    print(result.stdout)
    if result.stderr:
        print("ERRORS:\n", result.stderr)
finally:
    os.remove(ps_path)
>>>>>>> 701af31f89c54f74163107adb5acd0309ebdc4ba
