<div align="center">

## Compact Microsoft Access Database Through ADO


</div>

### Description

Although ADO specification does not provide objects to compact Microsoft Access databases, this capability can be achieved by using the ADO extension: Microsoft Jet OLE DB Provider and Replication Objects (JRO). This capability was implemented for the first time in the JET OLE DB Provider version 4.0 (Msjetoledb40.dll) and JRO version 2.1 (Msjro.dll). These DLL files are available after the install of MDAC 2.1. You can download the latest version of MDAC from the following Web site:
 
### More Info
 
the file you want to compact

No Side Effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Roni Saar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roni-saar.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/roni-saar-compact-microsoft-access-database-through-ado__1-45925/archive/master.zip)





### Source Code

```
Public Function CompactDatabase(strFileName As String) As Boolean
Dim objJro As jro.JetEngine
Dim objFileSystem As FileSystemObject
Dim strTmpFileName As String
 On Error GoTo EXIT_PROC
 Set objFileSystem = CreateObject("Scripting.FileSystemObject")
 strTmpFileName = objFileSystem.GetSpecialFolder(TemporaryFolder).Path & "\" & objFileSystem.GetFileName(strFileName)
 Set objJro = New jro.JetEngine
 objJro.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFileName, _
       "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTmpFileName & ";Jet OLEDB:Engine Type=5"
 objFileSystem.CopyFile strTmpFileName, strFileName
 objFileSystem.DeleteFile strTmpFileName, True
 CompactDatabase = True
EXIT_PROC:
 Set objFileSystem = Nothing
 Set objJro = Nothing
End Function
```

