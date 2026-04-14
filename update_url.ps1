$old = 'AKfycbwKdVVd4zRcGfPVXsm9l-cgQ8Cv6wbTtij1nB-24tZWKHHiwWk1ISQvXq_J5VI7ReIkdQ'
$new = 'AKfycbxelFju5HQlJWUaTsB-FT1mITgc2WnXX22h_LC7D8Ycy2fgFTXuethp8pUbsQThfqhYKg'

Get-ChildItem 'C:\Users\hashi\touka-tools\*.html' | ForEach-Object {
    $content = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    if ($content.Contains($old)) {
        $content = $content.Replace($old, $new)
        [System.IO.File]::WriteAllText($_.FullName, $content, [System.Text.Encoding]::UTF8)
        Write-Host "Updated: $($_.Name)"
    }
}

$testOcrPath = 'C:\Users\hashi\大久保先生\糖化\アンケート\Cloade Code\test-OCR\link_pdf.html'
if (Test-Path $testOcrPath) {
    $content = [System.IO.File]::ReadAllText($testOcrPath, [System.Text.Encoding]::UTF8)
    if ($content.Contains($old)) {
        $content = $content.Replace($old, $new)
        [System.IO.File]::WriteAllText($testOcrPath, $content, [System.Text.Encoding]::UTF8)
        Write-Host "Updated: test-OCR/link_pdf.html"
    }
}

Write-Host "Done."
