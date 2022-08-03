# converts a xlsx file with Name and Value Headers to a resx file
  function Convert-XlsxToResx
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)] [string]$XlsxFile,
    [Parameter(Mandatory = $true)] [string]$ResxFile
  )

  # read xlsx file Columns Name, Value
  $items = Import-Excel -Path $XlsxFile

  # resx writer
  Add-Type -AssemblyName System.Windows.Forms
  [System.Resources.ResXResourceWriter]$resxWriter = New-Object System.Resources.ResXResourceWriter ($ResxFile)

  # copy write values
  foreach ($item in $items)
  {
    $resxWriter.AddResource($item.Name,$item.Value);
  }

  # close writer
  $resxWriter.Close()
}

# converts a resx file to a xlsx file with Name and Value headers
function Convert-ResxToXlsx
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)] [string]$ResxFile,
    [Parameter(Mandatory = $true)] [string]$XlsxFile
  )

  # resx Reader
  Add-Type -AssemblyName System.Windows.Forms
  [System.Resources.ResXResourceReader]$resxReader = New-Object System.Resources.ResXResourceReader ($ResxFile)

  # read Name, Value lines
  $items = $resxReader | Select-Object Name,Value

  # export to excel
  $items | Export-Excel -Path $XlsxFile

  # close reader
  $resxReader.Close()
}


Export-ModuleMember -Function Convert-XlsxToResx
Export-ModuleMember -Function Convert-ResxToXlsx
