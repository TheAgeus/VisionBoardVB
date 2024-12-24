VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AgeusImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents img As Image
Attribute img.VB_VarHelpID = -1
Public index As Integer
Public parent As Form
Public title As String
Public descripcion As String
Public categoria As Integer
Public path As String
Public isNew As Byte
Public fecha As Date
Public change_pos_dimentions As Boolean

Public Function borrar(miSqlInstance As CSQLServer)
    With miSqlInstance.SQLCommand
        .Clear
        .Add "DELETE FROM propositos "
        .Add "where titulo = '" & title & "' "
    End With

    miSqlInstance.ExecSQL
End Function

Private Sub img_Click()
    ' Quitar el focus de la anterior
    If Not parent.currentImg Is Nothing Then
        parent.currentImg.img.BorderStyle = 0
    End If
    ' Simular que esta focuseado
    Me.img.BorderStyle = 1
    Set parent.currentImg = Me
End Sub

' Se guarda localmente la imagen con su index como nombre y se retorna el path db
' ya que si el proyecto esta en otro lado, no funcionarìa con mi path local
Private Function saveImgLocal()
    Dim local_path As String
    Dim db_path As String
    
    ' Path local
    local_path = App.path & "\imgs\" & title & ".jpg"
    
    ' Path a guardar en db
    db_path = "\imgs\" & title & ".jpg"
    
    ' Salvar imagen local
    FileCopy path, local_path
    
    saveImgLocal = db_path
End Function

' Se registran los datos del proposito en la base de datos
Public Function save(miSqlInstance As CSQLServer)

    ' Si es nueva la tarea o se cambio de pos
    If isNew = 1 Then
        isNew = 0
        ' Salvamos localmente la imagen y obtenemos el path
        Dim local_path As String
        local_path = saveImgLocal
    
        With miSqlInstance.SQLCommand
            .Clear
            .Add "INSERT INTO propositos "
            .Add "(fecha, titulo, descripcion, "
            .Add "categoria, local_path, is_new, "
            .Add "img_width, img_height, "
            .Add "img_top, img_left) "
            .Add "VALUES ('" & fecha & "', "
            .Add "'" & title & "',"
            .Add "'" & descripcion & "',"
            .Add "" & categoria & ","
            .Add "'" & local_path & "',"
            .Add "" & 0 & ","
            .Add "" & img.Width & ","
            .Add "" & img.Height & ","
            .Add "" & img.Top & ","
            .Add "" & img.Left & ""
            .Add ")"
        End With
    
        miSqlInstance.ExecSQL
    ElseIf change_pos_dimentions Then
        With miSqlInstance.SQLCommand
            .Clear
            .Add "UPDATE propositos SET "
            .Add "img_width = " & img.Width & ", img_height = " & img.Height & ","
            .Add "img_top = " & img.Top & ", img_left = " & img.Left & " "
            .Add "where titulo = '" & title & "' "
        End With
    
        miSqlInstance.ExecSQL
    End If

End Function

