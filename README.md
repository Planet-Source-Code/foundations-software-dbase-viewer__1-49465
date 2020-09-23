<div align="center">

## dBase Viewer


</div>

### Description

View the contents of a dBASE file (.DBF). Good example of how to use ListView and ADO together. Code can be easily modified to handle MSAccess files or even CSV files. Only thing particular about dBASE as opposed to other databases is the manner in which the Connection and SELECT strings are created.

oConn.Open "Driver={Microsoft dBASE Driver (*.dbf)};" & _

"DriverID=277;" & _

"Dbq=c:\somepath"

Then specify the filename in the SQL statement:

oRs.Open "Select * From user.dbf", oConn, , ,adCmdText

Hope someone finds this useful ... Cheers
 
### More Info
 
None that I know of but please report any bugs


<span>             |<span>
---                |---
**Submitted On**   |2003-10-27 02:43:00
**By**             |[Foundations Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/foundations-software.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[dBase\_View16635210272003\.zip](https://github.com/Planet-Source-Code/foundations-software-dbase-viewer__1-49465/archive/master.zip)








