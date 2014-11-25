# Function to get Process Owner
function Get-ProcessOwner {
   param (
      [String]$name,
      [Int]$id
   )

# Get processes using WMI
$p=get-wmiobject win32_process

# Now filter the list
if ($name) {$p=$p | where {$_.name -like "$name"}}
if ($id)   {$p=$p | where {$_.processid -eq $id}}

# Display the results
$p | select name, processid, @{n="owner";e={$_.getowner().user}} | format-table -autosize
}
