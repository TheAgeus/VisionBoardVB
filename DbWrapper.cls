VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public miSql As CSQLServer

Public Function read_txt(path_txt As String)

    Dim fileNumber As Integer
    Dim fileContent As String
    
    fileNumber = FreeFile
    
    ' Open the file for reading
    Open path_txt For Input As #fileNumber
    
    ' Read the entire file content
    fileContent = Input$(LOF(fileNumber), fileNumber) ' LOF returns the length of the file
    
    ' Display the file content
    read_txt = fileContent
    
    ' Close the file
    Close #fileNumber

End Function

Public Function ConnectDB() As Boolean

On Error GoTo ERR_HANDLER
    
    Set miSql = New CSQLServer
    
    miSql.ConnectionString = read_txt(App.path & "\connection_string.txt")
    
    ConnectDB = True
    miSql.Connect
    
    Exit Function

ERR_HANDLER:
    MsgBox "Error al conectase a la base de datos: " & Err.Description _
    , vbCritical, "Error"
    ConnectDB = False

End Function


' Example query
Public Function select100()

    With miSql.SQLCommand
        .Clear
        .Add "select top 100 * from test_table"

    End With

    miSql.ExecSQL

    If miSql.HasRecords Then
        Do While Not miSql.EOF
            Debug.Print (miSql.Value("texto"))
            miSql.MoveNext
        Loop
    End If

End Function

