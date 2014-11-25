function Unblock-File {
	[cmdletbinding()]
	param ([parameter(Mandatory=$true, ValueFromPipeline=$true)] [IO.FileInfo] $FilePath)
	begin {
		#http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
		Add-Type -Namespace PsFile -Name NtfsSecurity -MemberDefinition @"
			[DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool DeleteFile(string name);
			public static bool Unblock(string filePath) {
				return DeleteFile(filePath + ":Zone.Identifier");
			}
"@
	}
	process {
		try {
			#Discard the boolean result.
			[PsFile.NtfsSecurity]::Unblock($FilePath.FullName) > $null
		} catch {Write-Error (
				"Failed to unblock file '{0}'. The error was: '{1}'." -f $FilePath, $_)}
	}
	end {}
}
