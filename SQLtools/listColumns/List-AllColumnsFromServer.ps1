param(
[Parameter (
            Mandatory=$true, 
            Position = 1,
            HelpMessage = "Give Server name here, use '.' for localhost."
            )
] [String] $Servername = "." 
)

Invoke-Sqlcmd -HostName  $Servername -Query "sp_MSforeachdb 'select * from ?.information_schema.columns'" | Out-GridView