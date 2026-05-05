# Download catalog (already cached from last run, or re-downloads)
$cab = "$env:TEMP\DellCatalogPC.cab"
$xml = "$env:TEMP\CatalogPC.xml"
curl.exe --silent --location "https://downloads.dell.com/catalog/CatalogPC.cab" --output $cab
expand.exe $cab $xml

# Find all SoftwareComponent nodes that mention Latitude 7430 and print them
[xml]$cat = [System.IO.File]::ReadAllText($xml)
$cat.SelectNodes("//*[local-name()='SoftwareComponent']") | Where-Object {
    $_.OuterXml -match "(?i)Latitude.7430|F8X0F|0B0B"
} | Select-Object -First 3 | ForEach-Object { $_.OuterXml }