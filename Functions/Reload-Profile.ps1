# Function to reload profile
function Reload-Profile {
   @(
      $Profile.AllUsersAllHosts,
      $Profile.AllUsersCurrentHost,
      $Profile.CurrentUserAllHosts,
      $Profile.CurrentUserCurrentHost
   ) | % {
      if(Test-Path $_){
         Write-Verbose "Running $_"
         . $_
     }
   }
}
