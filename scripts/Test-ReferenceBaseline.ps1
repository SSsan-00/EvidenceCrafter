[CmdletBinding()]
param(
    [string]$ReferenceRoot = 'C:\work\Macro\Case&Evidence'
)

$ErrorActionPreference = 'Stop'
$expected = [ordered]@{
    'escape\BetaEvidenceGenerator.bas' = '7BC85DF208FA9A220714F9D40AE5FFCAF720F58CF8525C998ADC6E114AE11B51'
    'MacroModulesGuide.md' = 'CA8FE73561F6782904AD14481708E4A0A70BDC0592803C3064F46BAC726B0CC3'
    'MacroTest.xlsm' = '36A7C04CD0D6A50200A07728F691DF5DDE1566C3EC324EAC7E20AFE6A35EBAC8'
    'S00-000-00TESTX_【共通】02TEST2_単体テストエビデンス_初期開発.xlsx' = 'D2FAC5E708F461AC030D543D81EE88779D347F7C35369BB2E9AB430D038DEF8D'
    'S00-000-00TESTX_【個別】02TEST2_単体テストエビデンス_初期開発.xlsx' = '8C98F2E545DAFF407836E2842DEBE021B5CF24916FC71FFE99A67FD1755873CA'
    'S00-000-00TESTX_単体テストケース_初期開発_003.xlsx' = 'C10C6D120FAF51E1FB1484EBEC7DE52F343BB8FB9265C76C8FCFC85D135FF666'
}

$results = foreach ($relativePath in $expected.Keys) {
    $path = Join-Path $ReferenceRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Reference file is missing: $path"
    }

    $actual = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash
    [pscustomobject]@{
        File = $relativePath
        Unchanged = $actual -eq $expected[$relativePath]
        Expected = $expected[$relativePath]
        Actual = $actual
    }
}

$results | Format-Table File, Unchanged -AutoSize
if ($results.Unchanged -contains $false) {
    throw 'One or more Case&Evidence reference files changed.'
}
