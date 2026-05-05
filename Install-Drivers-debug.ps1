$xml = "$env:TEMP\CatalogPC.xml"
[xml]$cat = [System.IO.File]::ReadAllText($xml).TrimStart([char]0xFEFF)

# Find all nodes for 0B0B and show name + path
$cat.SelectNodes("//*[local-name()='SoftwareComponent']") | Where-Object {
    $_.SelectNodes(".//*[local-name()='Model']") | Where-Object { 
        $_.GetAttribute("systemID") -eq "0B0B" 
    }
} | ForEach-Object {
    $name = try { $_.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch { "?" }
    $path = $_.GetAttribute("path")
    Write-Host "NAME: $name"
    Write-Host "PATH: $path"
    Write-Host "---"
}