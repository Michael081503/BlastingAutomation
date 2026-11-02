param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [ValidateRange(1, 50)][int]$Cycles = 20
)

$ErrorActionPreference = 'Stop'
$retentionExcel = $null
$retentionWorkbook = $null
$retentionSheet = $null
$retentionRange = $null
try {
    $samples = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $retentionExcel = New-Object -ComObject Excel.Application
    $retentionExcel.Visible = $false
    $retentionExcel.DisplayAlerts = $false
    $retentionExcel.AskToUpdateLinks = $false
    Write-Output ('Testing Excel ' + $retentionExcel.Version)
    foreach ($sample in $samples) {
        for ($cycle = 1; $cycle -le $Cycles; $cycle++) {
            $retentionWorkbook = $retentionExcel.Workbooks.Open($sample.path, 0, $false)
            $retentionSheet = $retentionWorkbook.Worksheets.Item(1)
            $retentionRange = $retentionSheet.UsedRange
            $values = $retentionRange.Value2
            for ($row = 0; $row -lt $sample.expected.Count; $row++) {
                for ($field = 0; $field -lt $sample.columns.Count; $field++) {
                    $actual = [string]$values[($row + 2), [int]$sample.columns[$field]]
                    if ($actual -cne [string]$sample.expected[$row][$field]) {
                        throw ('Account digits changed in ' + [IO.Path]::GetFileName($sample.path) +
                               ', cycle ' + $cycle + ', row ' + ($row + 2) + ', column ' + $sample.columns[$field])
                    }
                }
            }
            $retentionWorkbook.Save()
            $retentionWorkbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($retentionRange)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($retentionSheet)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($retentionWorkbook)
            $retentionRange = $null
            $retentionSheet = $null
            $retentionWorkbook = $null
        }
        Write-Output ('PASS ' + [IO.Path]::GetFileName($sample.path) + ': ' + $Cycles + ' open/save/close cycles')
    }
} finally {
    if ($null -ne $retentionWorkbook) { $retentionWorkbook.Close($false) }
    foreach ($retentionObject in @($retentionRange, $retentionSheet, $retentionWorkbook)) {
        if ($null -ne $retentionObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($retentionObject) }
    }
    if ($null -ne $retentionExcel) {
        $retentionExcel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($retentionExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
