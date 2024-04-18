$date = (Get-Date).AddMonths(-3)
Get-GPO -All | Where-Object {
    ($_.CreationTime -ge $date) -or
    ($_.ModificationTime -ge $date)
} | Select-Object DisplayName, CreationTime, ModificationTime
