<div align="center">

## Deal with NULL in access


</div>

### Description

By default Access string fields contain NULL values unless a string value (including a blank string like "") has been assigned. When you read these fields using recordsets into VB string variables, you get a runtime type-mismatch error. here is a nice code to get rid of that
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dee Technologies](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dee-technologies.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dee-technologies-deal-with-null-in-access__1-30410/archive/master.zip)





### Source Code

```
Dim DB As Database
Dim RS As Recordset
Dim sYear As String
Set DB = OpenDatabase("Biblio.mdb")
Set RS = DB.OpenRecordset("Authors")
sYear = "" & RS![Year Born]
Just add a blank string to the returned field value.
```

