Sub RunMultipleQueries()

Dim dbs As DAO.Database

Set dbs = CurrentDb

' Execute runs both saved queries and SQL strings
cstrQueryName = "MakeTable Compare GTG Priority"
Debug.Print "Exporting: " & cstrQueryName
dbs.Execute cstrQueryName, dbFailOnError
DoEvents

Debug.Print "Finished"

Set dbs = Nothing

End Sub
