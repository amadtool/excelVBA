Attribute VB_Name = "Module1"
Sub Sheet2_Button1_Click()
    Dim objConnection As Object, objRecordset As Object
    Dim strDatabase As String
    Dim row As Integer, column As Integer
    Set objConnection = CreateObject("ADODB.Connection")
    Dim database_path As String, conn_string As String

    database_path = "C:\Users\Jesus.Ruiz\OneDrive - AECOM\Documents\test\db_sqlite\testDb.db"
    
    strDatabase = "Driver={SQLite3 ODBC Driver};Database=" & database_path

    Query = "drop table Table1; create table Table1 (Id, Name, Description, PRIMARY KEY('Name'));"
    row = 3
    column = 1
    While Not (IsNull(ActiveSheet.Cells(row, column).Value) Or IsEmpty(ActiveSheet.Cells(row, column).Value))
        Query = Query & "insert into FUND values('" & Cells(row, column) & "','" & Cells(row, column + 1) & "','" & Cells(row, column + 2) & "');"
        row = row + 1
    Wend

    objConnection.Open strDatabase
    objConnection.Execute Query
    objConnection.Close
End Sub