param (
  [Parameter(Mandatory=$True,Position=0)][string]$apikey,
  [Parameter(Mandatory=$True,Position=1)][string]$indir,
  [Parameter(Mandatory=$True,Position=2)][string]$outdir,
  [string]$format = "xlsx-single"
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName 'System.Net.Http'

$url = 'https://pdftables.com/api?format=' + $format + '&key=' + $apikey

$files = Get-ChildItem $indir -Recurse -File -Filter *.pdf |
Foreach-Object {
  $outsubdir = $outdir + $_.Directory.FullName.SubString($indir.Length)
  $outfile = $outdir + $_.FullName.SubString($indir.Length, $_.FullName.Length - $indir.Length - 3) + $format.Split('-')[0]
  New-Item -ItemType Directory -Force -Path $outsubdir
  Write-Host ('Converting "' + $_.FullName + '" to "' + $outfile + '"')

  Try {
    $client = New-Object System.Net.Http.HttpClient
    $request = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $url)
    
    $content = New-Object System.Net.Http.MultipartFormDataContent
    $fileStream = [System.IO.File]::OpenRead($_.FullName)
    $fileContent = New-Object System.Net.Http.StreamContent($fileStream)
    $content.Add($fileContent, 'unused', 'unused')

    $request.Content = $content 
    $result = $client.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    $stream = $result.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
    $fileStream = [System.IO.File]::Create($outfile)
    $stream.CopyTo($fileStream)
  }
  Catch {
    Write-Error $_
    exit 1
  }
  Finally {
    if ($fileContent -ne $null) { $fileContent.Dispose() }
    if ($fileStream -ne $null) { $fileStream.Dispose() }
    if ($content -ne $null) { $content.Dispose() }
    if ($request -ne $null) { $request.Dispose() }
    if ($client -ne $null) { $client.Dispose() }
  }
}
