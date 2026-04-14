$content = [System.IO.File]::ReadAllText("C:\Users\hashi\touka-tools\Code_gs_complete.js", [System.Text.Encoding]::UTF8)
Set-Clipboard -Value $content
Write-Host ("Copied " + $content.Length + " chars to clipboard")
