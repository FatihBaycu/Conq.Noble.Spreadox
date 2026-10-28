$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "Building Conq.Noble.Spreadox.sln (Release)..."
dotnet build .\Conq.Noble.Spreadox.sln -c Release
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Testing Conq.Noble.Spreadox.sln (Release) — default suite (excludes Category=Perf)..."
dotnet test .\Conq.Noble.Spreadox.sln -c Release --no-build --filter "Category!=Perf" --verbosity normal
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Packing NuGet to artifacts\ ..."
New-Item -ItemType Directory -Force -Path .\artifacts | Out-Null
dotnet pack .\src\Conq.Noble.Spreadox\Conq.Noble.Spreadox.csproj -c Release -o .\artifacts --no-build
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Done. Perf suite (optional): dotnet test .\Conq.Noble.Spreadox.sln -c Release --filter Category=Perf"
exit 0
