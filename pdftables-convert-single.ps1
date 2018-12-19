param (
  [Parameter(Mandatory=$True,Position=0)][string]$apikey,
  [Parameter(Mandatory=$True,Position=1)][string]$infile,
  [Parameter(Mandatory=$True,Position=2)][string]$outdir,
  [string]$format = "xlsx-single"
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName 'System.Net.Http'

$url = 'https://pdftables.com/api?format=' + $format + '&key=' + $apikey
$filename = ($infile.Split("\")[$infile.Split("\").count - 1]).Split(".")[0]
$outfile = $outdir + "\" + $filename + "." + $format.Split('-')[0]

Write-Host ('Converting "' + $infile + '" to "' + $outfile + '"')

Try {
  $client = New-Object System.Net.Http.HttpClient
  $request = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $url)
  
  $content = New-Object System.Net.Http.MultipartFormDataContent
  $fileStream = [System.IO.File]::OpenRead($infile)
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
