# Optional maintenance helpers. Dot-source alongside visio_helpers.ps1.

function Get-VisioMaintenanceShapeEntries {
  param([Parameter(Mandatory = $true)] $Shapes, [string] $ParentPath = '')
  foreach ($shape in $Shapes) {
    $path = if ($ParentPath) { $ParentPath + '/' + $shape.NameU } else { [string] $shape.NameU }
    [pscustomobject]@{ Shape = $shape; Path = $path }
    $children = $null
    try { if ($shape.Shapes.Count -gt 0) { $children = $shape.Shapes } } catch { }
    if ($null -ne $children) { Get-VisioMaintenanceShapeEntries -Shapes $children -ParentPath $path }
  }
}

function Get-VisioMaintenanceConnections {
  param([Parameter(Mandatory = $true)] $Shape)
  for ($i = 1; $i -le $Shape.Connects.Count; $i++) {
    $c = $Shape.Connects.Item($i)
    [pscustomobject]@{
      fromCell = [string] $c.FromCell.Name
      toShapeId = [int] $c.ToSheet.ID
      toNameU = [string] $c.ToSheet.NameU
      toCell = [string] $c.ToCell.Name
    }
  }
}

function Get-VisioNativeSnapshot {
  param([Parameter(Mandatory = $true)] $Document, [string[]] $PageNames = @())
  $pages = @()
  foreach ($page in $Document.Pages) {
    if ($PageNames.Count -gt 0 -and $page.NameU -notin $PageNames -and $page.Name -notin $PageNames) { continue }
    $shapes = @()
    foreach ($entry in @(Get-VisioMaintenanceShapeEntries -Shapes $page.Shapes)) {
      $s = $entry.Shape
      $cells = [ordered]@{}
      foreach ($name in @('PinX','PinY','Width','Height','Angle','BeginX','BeginY','EndX','EndY','ShapeRouteStyle','ConLineRouteExt','LineColor','LineWeight','FillForegnd','FillPattern','Char.Size','Char.Color','Char.Font','Para.HorzAlign','VerticalAlign','TxtPinX','TxtPinY')) {
        if ($s.CellExistsU($name,0)) { $cells[$name] = [ordered]@{ formulaU = [string] $s.CellsU($name).FormulaU; resultIU = [double] $s.CellsU($name).ResultIU } }
      }
      $links = @()
      foreach ($link in $s.Hyperlinks) { $links += [ordered]@{ address=[string]$link.Address; subAddress=[string]$link.SubAddress; description=[string]$link.Description } }
      $master = ''
      try { $master = [string] $s.Master.NameU } catch { }
      $shapes += [ordered]@{ id=[int]$s.ID; nameU=[string]$s.NameU; path=$entry.Path; text=[string]$s.Text; oneD=([int]$s.OneD -ne 0); master=$master; cells=$cells; links=$links; connections=@(Get-VisioMaintenanceConnections -Shape $s) }
    }
    $pages += [ordered]@{ id=[int]$page.ID; name=[string]$page.Name; nameU=[string]$page.NameU; width=[double]$page.PageSheet.CellsU('PageWidth').ResultIU; height=[double]$page.PageSheet.CellsU('PageHeight').ResultIU; shapes=$shapes }
  }
  if ($PageNames.Count -gt 0) {
    foreach ($name in $PageNames) { if (@($pages | Where-Object {$_.name -eq $name -or $_.nameU -eq $name}).Count -eq 0) { throw "Requested page not found: $name" } }
  }
  [pscustomobject]@{ path=[string]$Document.FullName; saved=[bool]$Document.Saved; capturedUtc=[DateTime]::UtcNow.ToString('o'); pages=$pages }
}

function Save-VisioDocumentBackup {
  param(
    [Parameter(Mandatory = $true)] $Document,
    [Parameter(Mandatory = $true)] [string] $BackupDirectory,
    [string[]] $RelatedPaths = @()
  )
  $source = [string] $Document.FullName
  if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw 'Backup requires a previously saved document with an existing file.' }
  $inputs = @($source)
  foreach ($path in $RelatedPaths) {
    $resolved = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw "Related file not found: $resolved" }
    if ($resolved -notin $inputs) { $inputs += $resolved }
  }
  $folder = Join-Path ([IO.Path]::GetFullPath($BackupDirectory)) ([IO.Path]::GetFileNameWithoutExtension($source) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
  $null = $Document.Save()
  $files = @()
  for ($i=0; $i -lt $inputs.Count; $i++) {
    $destination = Join-Path $folder (('{0:D2}_' -f $i) + [IO.Path]::GetFileName($inputs[$i]))
    Copy-Item -LiteralPath $inputs[$i] -Destination $destination
    $files += [ordered]@{ source=$inputs[$i]; backup=$destination; sha256=(Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash }
  }
  $report = [pscustomobject]@{ directory=$folder; documentSaved=[bool]$Document.Saved; files=$files }
  $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $folder 'backup.json') -Encoding utf8
  return $report
}

function Export-VisioPagePreview {
  param(
    [Parameter(Mandatory = $true)] $Page,
    [Parameter(Mandatory = $true)] [string] $Path,
    [ValidateRange(1,1200)] [int] $PixelsPerInch = 144,
    [int] $PixelWidth = 0,
    [int] $PixelHeight = 0
  )
  $resolved = [IO.Path]::GetFullPath($Path)
  if ([IO.Path]::GetExtension($resolved) -ne '.png') { throw 'This preview helper supports PNG only.' }
  if (($PixelWidth -eq 0) -ne ($PixelHeight -eq 0) -or $PixelWidth -lt 0 -or $PixelHeight -lt 0) { throw 'Specify both positive PixelWidth and PixelHeight, or omit both.' }
  if ($PixelWidth -eq 0) {
    $PixelWidth = [int][Math]::Max(1,[Math]::Round($Page.PageSheet.CellsU('PageWidth').ResultIU*$PixelsPerInch))
    $PixelHeight = [int][Math]::Max(1,[Math]::Round($Page.PageSheet.CellsU('PageHeight').ResultIU*$PixelsPerInch))
  }
  New-Item -ItemType Directory -Path (Split-Path -Parent $resolved) -Force | Out-Null
  $app = $Page.Application
  $oldSize=0; $oldWidth=0.0; $oldHeight=0.0; $oldUnits=0
  $null = $app.Settings.GetRasterExportSize([ref]$oldSize,[ref]$oldWidth,[ref]$oldHeight,[ref]$oldUnits)
  try {
    $null = $app.Settings.SetRasterExportSize(3,$PixelWidth,$PixelHeight,0)
    $null = $Page.Export($resolved)
  } finally {
    $null = $app.Settings.SetRasterExportSize($oldSize,$oldWidth,$oldHeight,$oldUnits)
  }
  $file=Get-Item -LiteralPath $resolved
  [pscustomobject]@{ page=[string]$Page.NameU; path=$resolved; requestedPixelWidth=$PixelWidth; requestedPixelHeight=$PixelHeight; bytes=$file.Length }
}

function Get-VisioPageConnectorReport {
  param([Parameter(Mandatory = $true)] $Page, [object[]] $ExpectedEdges = @())
  $entries = @(Get-VisioMaintenanceShapeEntries -Shapes $Page.Shapes)
  $details = @()
  foreach ($entry in $entries) {
    $s=$entry.Shape
    if ([int]$s.OneD -eq 0) { continue }
    $connections=@(Get-VisioMaintenanceConnections -Shape $s)
    $begin=@($connections | Where-Object {$_.fromCell -match '^Begin[XY]$'})
    $end=@($connections | Where-Object {$_.fromCell -match '^End[XY]$'})
    $master=''; try{$master=[string]$s.Master.NameU}catch{}
    $route=$null; if($s.CellExistsU('ConLineRouteExt',0)){$route=[int]$s.CellsU('ConLineRouteExt').ResultIU}
    $details += [pscustomobject]@{ id=[int]$s.ID; nameU=[string]$s.NameU; path=$entry.Path; master=$master; routeExtension=$route; begin=$begin; end=$end; beginFormula=[string]$s.CellsU('BeginX').FormulaU; endFormula=[string]$s.CellsU('EndX').FormulaU; endpointsGlued=($begin.Count -gt 0 -and $end.Count -gt 0) }
  }
  $expectedChecks=@(); $seen=@{}
  foreach($edge in $ExpectedEdges){
    if(-not $edge.id){throw 'Each expected edge needs its native NameU in id.'}
    if($seen.ContainsKey([string]$edge.id)){throw "Duplicate expected edge id: $($edge.id)"}
    $seen[[string]$edge.id]=$true
    $issues=@(); $matches=@($details | Where-Object {$_.nameU -eq $edge.id})
    if($matches.Count -ne 1){$issues += 'Expected connector absent or ambiguous'}
    else{
      $connector=$matches[0]
      if(-not $connector.endpointsGlued){$issues += 'A connector endpoint is not glued'}
      foreach($side in @('from','to')){
        $idProperty=$edge.PSObject.Properties[$side+'Id']
        if($null -ne $idProperty){$targets=@($entries | Where-Object {[int]$_.Shape.ID -eq [int]$idProperty.Value})}
        else{
          $nameProperty=$edge.PSObject.Properties[$side]
          if($null -eq $nameProperty -or -not $nameProperty.Value){throw "Expected edge $($edge.id) needs $side or $($side)Id"}
          $targets=@($entries | Where-Object {$_.Shape.NameU -eq [string]$nameProperty.Value})
        }
        if($targets.Count -ne 1){$issues += "Expected $side node absent or ambiguous"; continue}
        $records=if($side -eq 'from'){@($connector.begin)}else{@($connector.end)}
        $wanted=[int]$targets[0].Shape.ID
        if($records.Count -eq 0 -or @($records | Where-Object {$_.toShapeId -ne $wanted}).Count -gt 0){$issues += "Wrong $side endpoint target"}
      }
      if($edge.connector -and $connector.master -ne [string]$edge.connector){$issues += 'Unexpected connector master'}
    }
    $expectedChecks += [pscustomobject]@{ id=[string]$edge.id; passed=($issues.Count -eq 0); issues=$issues }
  }
  [pscustomobject]@{
    page=[string]$Page.NameU
    connectorCount=$details.Count
    allEndpointsGlued=(@($details | Where-Object {-not $_.endpointsGlued}).Count -eq 0)
    expectedEdgesChecked=$ExpectedEdges.Count
    expectedTargetsVerified=($ExpectedEdges.Count -gt 0 -and @($expectedChecks | Where-Object {-not $_.passed}).Count -eq 0)
    connectors=$details
    expectedChecks=$expectedChecks
  }
}
