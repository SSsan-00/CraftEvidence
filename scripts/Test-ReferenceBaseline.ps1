[CmdletBinding()]
param(
    [string]$ReferenceRoot = 'C:\work\Macro\Case&Evidence'
)

$ErrorActionPreference = 'Stop'
$expected = [ordered]@{
    'escape\BetaEvidenceGenerator.bas' = '0C1B0F5D8108D74D738BB9BE9C34893A37F407B5634FE9BCF199BE57B74ECCF5'
    'MacroModulesGuide.md' = '25319391FF27B63412B78CF6A2ED66CF2FDC18003436FBE8F706D2FA201D9B2B'
    'MacroTest.xlsm' = '72618E3D34ED63EB716FCE7508A1F2D5CE93A6FBA7D03A46A6A346CBA45E4A3F'
    'S00-000-00TESTX_【共通】02TEST2_単体テストエビデンス_初期開発.xlsx' = 'D2FAC5E708F461AC030D543D81EE88779D347F7C35369BB2E9AB430D038DEF8D'
    'S00-000-00TESTX_【個別】02TEST2_単体テストエビデンス_初期開発.xlsx' = 'F105816F85476CD58399A24A39706F1D172BA9E380EC6D512749C18F52DFDB52'
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
