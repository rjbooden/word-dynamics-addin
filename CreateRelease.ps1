# Build front-end
cd .\word-dynamics-addin
npm run update-addin
cd ..\word-dynamics-api
dotnet publish --configuration Release
cd ..

$publishPath = ".\word-dynamics-api\bin\Release\net7.0\publish"

$manifest = Get-Content -Path "$publishPath\wwwroot\manifest.xml"
[xml]$manifestXml = $manifest
$version = $manifestXml.GetElementsByTagName("Version")[0].InnerText

$manifest = $manifest -replace "word-dynamics-addin.azurewebsites.net", "[your_url]"
$manifest = $manifest -replace "3830ed35-f717-46b8-acfe-2c6bb1c6cc95", "[your clientid]"

Set-Content -Path "$publishPath\wwwroot\manifest.xml" -Value $manifest -Force

$compress = @{ 
    Path = @()
    CompressionLevel = "Optimal"
    DestinationPath = ".\release\release_$version.zip"
}

$items = Get-ItemProperty -Path "$publishPath\*"
$items | ForEach-Object {
    if ($_.Name -ine "appsettings.Development.json") {
        $compress.Path += , $_.FullName
    }
}

Compress-Archive @compress -Force
