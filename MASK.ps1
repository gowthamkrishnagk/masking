<#
.SYNOPSIS
    Masks a selected column (or columns) in a CSV file and saves the result
    to a new copy, leaving the original file untouched.

.DESCRIPTION
    1. Asks for the path to a CSV file.
    2. Reads it and lists all column names.
    3. Asks which column(s) to mask (by number).
    4. For each selected column, masking is PARTIAL (not everything is
       hidden), and every masked value is guaranteed UNIQUE within that
       column -- two different rows can never end up with the same masked
       value:
         - Numeric column (phone, SSN, etc. -- values that parse as plain
           numbers): the last few digits stay visible, the rest are
           replaced with 0s, e.g. "9876543210" -> "0000013210". The
           hidden digits are not simply all zero -- they encode a number
           based on the row's position, so no two rows can collide.
         - Mixed column (email, name, or anything with letters and
           punctuation, e.g. a dashed SSN "123-45-6789"): the first
           letter/digit of each "word" (separated by spaces, dots,
           dashes, @, etc.) stays visible, the rest is replaced with X,
           e.g. "Alice Johnson" -> "AXXXX JXXXXX1", and
           "alice.johnson@example.com" -> "aXXXX.jXXXXX1@example.com"
           (the domain is always left fully visible). Punctuation like
           . - @ is never touched, so the format still looks recognizable.
    5. Saves the masked data as "<originalname>_masked.csv" in the same
       folder as the original.

.NOTES
    Uniqueness relies on each row getting a distinct 1-based row number
    encoded into the hidden characters. As long as the number of rows in
    your CSV doesn't exceed the "capacity" of the hidden portion (e.g. 6
    hidden digits = 1,000,000 possible combinations), every masked value
    in that column stays unique. This is comfortably true for
    normal-sized CSV files.

    Run it by right-clicking the file -> "Run with PowerShell", or from a
    PowerShell prompt:
        powershell -ExecutionPolicy Bypass -File .\Mask-CsvColumn.ps1
#>

# ---- Helper: is a value "numeric" in the practical sense? ----
# A value counts as numeric if, after stripping common formatting characters
# (spaces, dashes, slashes, parentheses, plus signs, decimal points), only
# digits are left. This correctly treats formatted identifiers like a dashed
# SSN ("123-45-6789"), a parenthesized phone number, or a date in either
# "2024-01-15" or "01/15/2024" style as numeric, while still treating
# anything with actual letters (emails, names) as mixed.
function Test-IsNumeric {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $true }
    $stripped = $Value -replace '[\s\-\/\(\)\+\.]', ''
    if ([string]::IsNullOrEmpty($stripped)) { return $false }
    return $stripped -match '^\d+$'
}

# ---- Helper: encode a number as a FULL-WIDTH, zero-padded string in a given base ----
# Used for numeric columns -- e.g. row 13 with 6 hidden digits -> "000013".
# This is a plain fixed-width positional numeral, so every number in range
# maps to one unique string; that's what keeps masked numbers collision-free.
function Convert-ToBase {
    param(
        [long]$Number,
        [string]$Alphabet,
        [int]$Width
    )
    if ($Width -le 0) { return '' }
    $base     = $Alphabet.Length
    $capacity = [long][Math]::Pow($base, $Width)
    $n = $Number % $capacity
    $result = @('0') * $Width
    for ($pos = $Width - 1; $pos -ge 0; $pos--) {
        $result[$pos] = $Alphabet[$n % $base]
        $n = [long]([Math]::Floor($n / $base))
    }
    return -join $result
}

# ---- Helper: minimal (no leading-zero) representation of a number in a given base ----
function Convert-ToMinimalBase {
    param(
        [long]$Number,
        [string]$Alphabet
    )
    $base = $Alphabet.Length
    if ($Number -eq 0) { return [string]$Alphabet[0] }
    $digits = New-Object System.Collections.Generic.List[char]
    $n = $Number
    while ($n -gt 0) {
        $digits.Insert(0, $Alphabet[$n % $base])
        $n = [long]([Math]::Floor($n / $base))
    }
    return -join $digits
}

# ---- Helper: mostly $FillChar, with a short unique code right-aligned at the end ----
# Used for mixed/X-masked columns, e.g. width 5 with $FillChar 'X' -> "XXXA1".
# $Alphabet must NOT contain $FillChar, or the pad/code boundary becomes
# ambiguous and two different numbers could produce the same string.
function Get-FillerEncoded {
    param(
        [long]$Number,
        [string]$Alphabet,
        [char]$FillChar,
        [int]$Width
    )
    if ($Width -le 0) { return '' }
    $base     = $Alphabet.Length
    $capacity = [long][Math]::Pow($base, $Width)
    $n = $Number % $capacity
    $digits = Convert-ToMinimalBase -Number $n -Alphabet $Alphabet
    if ($digits.Length -gt $Width) { $digits = $digits.Substring($digits.Length - $Width) }
    $padCount = $Width - $digits.Length
    return ([string]$FillChar * $padCount) + $digits
}

# ---- Helper: mask a numeric value -> last few digits visible, rest becomes 0s ----
function Get-MaskedNumeric {
    param(
        [string]$Value,
        [int]$RowNumber,
        [int]$VisibleDigits = 4
    )
    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    $chars = $Value.ToCharArray()
    $digitPositions = @()
    for ($i = 0; $i -lt $chars.Length; $i++) {
        if ($chars[$i] -match '\d') { $digitPositions += $i }
    }

    $totalDigits = $digitPositions.Count
    if ($totalDigits -eq 0) { return $Value }

    # Always keep at least 1 digit hidden; otherwise keep up to $VisibleDigits visible
    $visibleCount = [Math]::Min($VisibleDigits, [Math]::Max(0, $totalDigits - 1))
    $hiddenCount  = $totalDigits - $visibleCount

    if ($hiddenCount -gt 0) {
        $hiddenPositions = $digitPositions[0..($hiddenCount - 1)]
        $encoded = Convert-ToBase -Number $RowNumber -Alphabet '0123456789' -Width $hiddenCount
        for ($k = 0; $k -lt $hiddenPositions.Count; $k++) {
            $chars[$hiddenPositions[$k]] = $encoded[$k]
        }
    }

    return -join $chars
}

# ---- Helper: mask alphanumeric text -> first char of each "word" visible, rest becomes X ----
# A "word" is a run of letters/digits; anything else (space, ., -, @, ...) is
# left untouched and also resets what counts as the start of the next word.
function Get-MaskedAlnum {
    param(
        [string]$Text,
        [int]$RowNumber
    )
    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    # 'X' is deliberately excluded from the encoding alphabet -- it is reserved
    # purely as filler, which is what makes the fill/code boundary unambiguous.
    $alphabetNoX = '0123456789ABCDEFGHIJKLMNOPQRSTUVWYZ'

    $chars = $Text.ToCharArray()
    $hiddenPositions = @()
    $atWordStart = $true
    for ($i = 0; $i -lt $chars.Length; $i++) {
        if ($chars[$i] -notmatch '[A-Za-z0-9]') {
            $atWordStart = $true
            continue
        }
        if ($atWordStart) {
            $atWordStart = $false   # first letter/digit of a word -> stays visible
        } else {
            $hiddenPositions += $i
        }
    }

    if ($hiddenPositions.Count -gt 0) {
        $encoded = Get-FillerEncoded -Number $RowNumber -Alphabet $alphabetNoX -FillChar 'X' -Width $hiddenPositions.Count
        for ($k = 0; $k -lt $hiddenPositions.Count; $k++) {
            $chars[$hiddenPositions[$k]] = $encoded[$k]
        }
    }

    return -join $chars
}

# ---- Helper: mask a mixed value (handles email specially: domain stays fully visible) ----
function Get-MaskedMixed {
    param(
        [string]$Value,
        [int]$RowNumber
    )
    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    if ($Value -match '^([^@]+)@([^@]+)$') {
        $local  = $Matches[1]
        $domain = $Matches[2]
        $maskedLocal = Get-MaskedAlnum -Text $local -RowNumber $RowNumber
        return "$maskedLocal@$domain"
    }

    return Get-MaskedAlnum -Text $Value -RowNumber $RowNumber
}

# ---- 1. Ask for the CSV path ----
$csvPath = Read-Host "Enter the full path to the CSV file"

if (-not (Test-Path -LiteralPath $csvPath)) {
    Write-Host "File not found: $csvPath" -ForegroundColor Red
    exit 1
}

# ---- 2. Read the CSV ----
try {
    $data = Import-Csv -LiteralPath $csvPath
} catch {
    Write-Host "Failed to read CSV: $_" -ForegroundColor Red
    exit 1
}

if (-not $data -or $data.Count -eq 0) {
    Write-Host "The CSV file appears to be empty." -ForegroundColor Red
    exit 1
}

$columns = $data[0].PSObject.Properties.Name

Write-Host "`nColumns found in the CSV:"
for ($i = 0; $i -lt $columns.Count; $i++) {
    Write-Host ("  [{0}] {1}" -f ($i + 1), $columns[$i])
}

# ---- 3. Ask which column(s) to mask ----
$selection = Read-Host "`nEnter the number of the column to mask (or comma-separated numbers for several, e.g. 2,4)"
$selectedIndexes = $selection -split ',' |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -match '^\d+$' } |
    ForEach-Object { [int]$_ - 1 }

$selectedColumns = @()
foreach ($idx in $selectedIndexes) {
    if ($idx -ge 0 -and $idx -lt $columns.Count) {
        $selectedColumns += $columns[$idx]
    }
}

if ($selectedColumns.Count -eq 0) {
    Write-Host "No valid column selected. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host ("`nMasking column(s): {0}" -f ($selectedColumns -join ', '))

# ---- 4. Mask each selected column ----
foreach ($col in $selectedColumns) {
    # Decide numeric vs. combined by checking every non-blank value in the column
    $values = $data | ForEach-Object { $_.$col } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $isNumericColumn = $true
    foreach ($v in $values) {
        if (-not (Test-IsNumeric $v)) { $isNumericColumn = $false; break }
    }

    for ($i = 0; $i -lt $data.Count; $i++) {
        $rowNumber = $i + 1   # 1-based, strictly unique per row -> guarantees no masked duplicates
        $row = $data[$i]
        if ($isNumericColumn) {
            $row.$col = Get-MaskedNumeric -Value $row.$col -RowNumber $rowNumber
        } else {
            $row.$col = Get-MaskedMixed -Value $row.$col -RowNumber $rowNumber
        }
    }

    $kind = if ($isNumericColumn) { "numeric -> last digits visible, rest masked with 0s" } else { "combined -> first letter of each word visible, rest masked with X" }
    Write-Host ("  '{0}' detected as {1}" -f $col, $kind)
}

# ---- 5. Save as a new copy ----
$directory  = Split-Path -LiteralPath $csvPath -Parent
if ([string]::IsNullOrEmpty($directory)) { $directory = "." }
$baseName   = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
$extension  = [System.IO.Path]::GetExtension($csvPath)
$outputPath = Join-Path $directory ("{0}_masked{1}" -f $baseName, $extension)

$data | Export-Csv -LiteralPath $outputPath -NoTypeInformation

Write-Host "`nDone. Masked file saved to: $outputPath" -ForegroundColor Green