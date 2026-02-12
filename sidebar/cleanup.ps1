$path = "c:\Dev\Outlook_Sidebar\sidebar_main.py"
$lines = Get-Content $path
if ($lines.Count -lt 3100) {
    Write-Host "File already cleaned or too small. Current lines: $($lines.Count)"
    exit
}

# Keep Header (Imports)
$header = $lines[0..36]

# Keep SidebarWindow and Main
# Index 3004 is 'class SidebarWindow'
$body = $lines[3004..($lines.Count-1)]

$newContent = $header + $body
$newContent | Set-Content $path -Encoding UTF8
Write-Host "Cleanup complete. New line count: $($newContent.Count)"
