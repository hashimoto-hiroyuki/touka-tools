$old1 = 'https://script.google.com/macros/s/AKfycbzQCrKRX7nJgryTPsP2Aceh4_Ofyef2Ez2iBmHUGBYF3K15XYZk-5Na8XDIlLCqlAGtVQ/exec'
$new = 'https://script.google.com/a/macros/devine.co.jp/s/AKfycbxelFju5HQlJWUaTsB-FT1mITgc2WnXX22h_LC7D8Ycy2fgFTXuethp8pUbsQThfqhYKg/exec'

$basePath = 'C:\Users\hashi\大久保先生\糖化\アンケート\Cloade Code\test-OCR'

Get-ChildItem -Path $basePath -Filter '*.html' -Recurse | ForEach-Object {
    $content = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    if ($content.Contains($old1)) {
        $content = $content.Replace($old1, $new)
        [System.IO.File]::WriteAllText($_.FullName, $content, [System.Text.Encoding]::UTF8)
        Write-Host "Updated: $($_.FullName)"
    }
}

Write-Host "Done."
