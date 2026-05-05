$xml = "$env:TEMP\CatalogPC.xml"
[xml]$cat = [System.IO.File]::ReadAllText($xml).TrimStart([char]0xFEFF)

# How many SoftwareComponent nodes does XPath actually find?
$all = $cat.SelectNodes("//*[local-name()='SoftwareComponent']")
Write-Host "Total SoftwareComponent nodes: $($all.Count)"

# How many have 'driver pack' in the name?
$packs = $all | Where-Object {
    try { $_.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText -match "(?i)driver\s*pack" } catch { $false }
}
Write-Host "Driver pack nodes: $($packs.Count)"

# Of those, which ones have a Model with systemID 0B0B?
$packs | ForEach-Object {
    $name = try { $_.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch { "?" }
    $ids = $_.SelectNodes(".//*[local-name()='Model']") | ForEach-Object { $_.GetAttribute("systemID") }
    Write-Host "PACK: $name"
    Write-Host "  systemIDs: $($ids -join ', ')"
}