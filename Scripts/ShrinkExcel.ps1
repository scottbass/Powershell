<# HEADER
/*=====================================================================
Program Name            : ShrinkExcel.ps1
Purpose                 : Invoke Excel in "batch mode" as a COM
                          application to shrink ODS Excel Tagset
                          XML output
Powershell Version:     : v2.0
Input Data              : Excel XML file created by ODS Excel Tagsets
Output Data             : Native Excel workbook (default .xlsb format)

Originally Written by   : Scott Bass
Date                    : 06NOV2013
Program Version #       : 1.0

=======================================================================

Modification History    :

=====================================================================*/
#>

<#
.SYNOPSIS
Shrink Excel XML files.

.DESCRIPTION
Invoke Excel in "batch mode" as a COM application to shrink ODS Excel Tagset output.

.PARAMETER  xlFile
ODS Excel Tagset XML file

.PARAMETER  Path
Alias for xlFile

.PARAMETER  xlFormat
Excel output format (xlsx, xlsm, xlsb, xls).  Default is xlsb.

.PARAMETER  Format
Alias for xlFile

.EXAMPLE
.\ShrinkExcel.ps1 C:\Temp\Excel_Tagset.xml, or
.\ShrinkExcel.ps1 -xlFile C:\Temp\Excel_Tagset.xml, or
.\ShrinkExcel.ps1 -Path C:\Temp\Excel_Tagset.xml

Description
-----------
Shrink the ODS Excel Tagset XML file C:\Temp\Excel_Tagset.xml,
creating the default C:\Temp\Excel_Tagset.xlsb output file.

.EXAMPLE
.\ShrinkExcel.ps1 C:\Temp\Excel_Tagset.xml xlsx, or
.\ShrinkExcel.ps1 -xlFile C:\Temp\Excel_Tagset.xml -xlFormat xlsx, or
.\ShrinkExcel.ps1 -Path C:\Temp\Excel_Tagset.xml -Format xlsx

Description
-----------
Shrink the ODS Excel Tagset XML file C:\Temp\Excel_Tagset.xml,
creating the C:\Temp\Excel_Tagset.xlsx output file.

.INPUTS
Excel XML file created by ODS Excel Tagsets

.OUTPUTS
Native Excel workbook (default .xlsb format)

.NOTES
The output file is written to the same path and name as the input file,
but with the relevant xlsb, xlsx, xlsm, or xls extension.

This script also autofits all columns in all worksheets, setting the column
width based on the length of the longest variable in the column.

.LINK
http://technet.microsoft.com/en-us/library/ff730962.aspx
http://gallery.technet.microsoft.com/office/1d187163-8c81-4f1e-b6c9-bd8a41a680a7

#>

#region Parameters
param(
   # full path to the original Excel file (usually in XML format)
   [Parameter(
      Position=0,
      Mandatory=$true
   )]
   [Alias("Path")]
   [ValidateScript({Test-Path $_})]
   [System.IO.FileInfo]$xlFile
   ,
   # output format
   [Parameter(
      Position=1
   )]
   [Alias("Format")]
   [ValidateSet("xlsx","xlsm","xlsb","xls")]
   [String]$xlFormat="xlsb"
)
#endregion

#region functions
# See http://technet.microsoft.com/en-us/library/ff730962.aspx
# or Google "ReleaseComObject Excel"
function release-comobject($ref) {
   while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) -gt 0) {}
   [System.GC]::Collect()
   [System.GC]::WaitForPendingFinalizers()
}
#endregion

#region Main

# Keep running if an error occurs, usually to ensure Excel is closed down properly
$ErrorActionPreference="Continue"

# Get the directory path and basename from the source file
$xlPath=$xlFile.DirectoryName
$xlBasename=$xlFile.Basename

# Create enumerations for the various SaveAs formats
# See http://www.rondebruin.nl/mac/mac020.htm
switch ($xlFormat) {
   "xlsx" {$xlFormatNum=51; break}
   "xlsm" {$xlFormatNum=52; break}
   "xlsb" {$xlFormatNum=50; break}
   "xls"  {$xlFormatNum=56; break}
}

try {
   # Instantiate Excel as a COM application
   $Excel = New-Object -ComObject excel.application

   # Run in "batch mode"
   $Excel.visible = $False
   $Excel.displayalerts = $False

   # Open the XML workbook
   $Workbook = $Excel.Workbooks.Open($xlFile)

   #=============================================================================#
   # autofit row/column                                                          #
   # before we save the file, call AutoFit to improve the appearance             #
   #=============================================================================#

   foreach ($Worksheet in $Workbook.Worksheets) {
      [Void] $Worksheet.Activate()
      [Void] ($Worksheet.UsedRange).EntireColumn.AutoFit()
      [Void] ($Worksheet.UsedRange).Rows.AutoFit()

      # equivalent of Cntl-Home if freeze panes is active
      $row=$Excel.ActiveWindow.SplitRow + 1
      $col=$Excel.ActiveWindow.SplitColumn + 1
      [Void] ($Worksheet.Cells.Item($row,$col)).Select()

      release-comobject $Worksheet
   }

   # Select the first worksheet
   [Void] $Workbook.Sheets.Item(1).Activate()

   # Save to the same path and filename as the source file, but with a different extension
   [Void] $Workbook.SaveAs((Join-Path $xlPath "$xlBasename.$xlFormat"), $xlFormatNum)

   # Close the workbook
   [Void] $Workbook.Close()

   # Quit the Excel COM application
   [Void] $Excel.Quit()
}
catch {
   Write-Error $_
}
# Finally executes whether there is an error or not
finally {
   release-comobject $Workbook
   release-comobject $Excel
}
#endregion
