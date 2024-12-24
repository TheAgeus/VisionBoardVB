VERSION 5.00
Begin VB.Form VisionBoard 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VisionBoard"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "VisionBoard.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   480
      Left            =   5160
      TabIndex        =   3
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   480
      Left            =   3720
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   480
      Left            =   2160
      TabIndex        =   1
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAgregarActividad 
      Caption         =   "Agregar Actividad"
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblWArriba_ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"VisionBoard.frx":9965
      Height          =   780
      Left            =   8880
      TabIndex        =   6
      Top             =   7680
      Width           =   1050
   End
   Begin VB.Label lblYReducir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"VisionBoard.frx":999E
      Height          =   780
      Left            =   6720
      TabIndex        =   5
      Top             =   7680
      Width           =   1770
   End
   Begin VB.Label lblWArriba 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6720
      TabIndex        =   4
      Top             =   7440
      Width           =   45
   End
End
Attribute VB_Name = "VisionBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public imgsPath As String
Public currentImg As AgeusImg
Public countImages As Integer
Public imgWrapper As Variant
Public gSqlConnection As DbWrapper
Public categorias As Variant

Private Function loadCategorias()
    
    With gSqlConnection.miSql.SQLCommand
        .Clear
        .Add "select * from categorias_proposito with(nolock)"
    End With
    
    gSqlConnection.miSql.ExecSQL
    
    ReDim categorias(0, 1) ' Initial size

    Do While Not gSqlConnection.miSql.EOF
        ' Create a temporary array with 1 more row
        Dim tempCategorias() As Variant
        ReDim tempCategorias(UBound(categorias) + 1, 1)
        
        ' Copy the old array data into the new array
        Dim i As Integer
        For i = 0 To UBound(categorias)
            tempCategorias(i, 0) = categorias(i, 0)
            tempCategorias(i, 1) = categorias(i, 1)
        Next i
        
        ' Now add the new data to the last position
        tempCategorias(UBound(tempCategorias), 0) = gSqlConnection.miSql.Recordset.Fields(0)
        tempCategorias(UBound(tempCategorias), 1) = gSqlConnection.miSql.Recordset.Fields(1)
        
        ' Assign tempCategorias back to categorias
        categorias = tempCategorias
        
        ' Move to the next record
        gSqlConnection.miSql.MoveNext
    Loop
    
End Function

Private Sub cmdBorrar_Click()
    If currentImg.img Is Nothing Then
        MsgBox "Seleccione una imagen primero"
        Exit Sub
    End If
    
    respuesta = MsgBox("¿Está seguro de que desea borrar?", vbQuestion + vbYesNo, "Confirmar Borrado")

    ' Comprobar la respuesta del usuario
    If respuesta = vbYes Then
        ' Borrar de db
        currentImg.borrar gSqlConnection.miSql
        
        ' Borrar local img
        currentImg.img.Visible = False
        Kill currentImg.path
        
        ' Borrar local
        ' Search for the matching object
        Dim indexToRemove As Integer
        Dim found As Boolean
        
        For i = 1 To UBound(imgWrapper)
            If imgWrapper(i).title = currentImg.title Then
                indexToRemove = i
                found = True
                Exit For  ' Exit loop once the item is found
            End If
        Next i
    
        ' If we found the object to remove
        If found Then
            ' Shift elements to the left starting from the found index
            For i = indexToRemove To UBound(imgWrapper) - 1
                Set imgWrapper(i) = imgWrapper(i + 1)
            Next i
            
            ' Set the last element to Nothing (optional)
            Set imgWrapper(UBound(imgWrapper)) = Nothing
            
            ' Resize the array to remove the last element
            ReDim Preserve imgWrapper(LBound(imgWrapper) To UBound(imgWrapper) - 1)
        End If
        
    Else
        ' El usuario ha pulsado "No", no hacer nada
        MsgBox "Operación cancelada."
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    ' Guardar las posiciones y las rutas de las imagenes nuevas
    If UBound(imgWrapper) > 0 Then
        Dim i As Integer
        For i = 1 To UBound(imgWrapper)
            imgWrapper(i).save gSqlConnection.miSql
        Next i
        MsgBox "Guardado correcto"
    Else
        MsgBox "No hay propositos a guardar"
    End If
End Sub

Private Function getCategoriaByIndex(dbIndex As Integer)
    
    If UBound(categorias) > 0 Then
        Dim i As Integer
        For i = 1 To UBound(categorias)
            If categorias(i, 0) = dbIndex Then
                getCategoriaByIndex = categorias(i, 1)
                Exit Function
            End If
        Next i
    Else
    getCategoriaByIndex = "no encontrada"
End Function

Private Sub cmdMostrar_Click()

    If currentImg.img Is Nothing Then
        MsgBox "Seleccione una imagen primero"
        Exit Sub
    End If

    Detalle.txtTitulo = currentImg.title
    Detalle.txtCategoria = currentImg.categoria
    Detalle.txtFecha = currentImg.fecha
    Detalle.txtDesc = currentImg.descripcion
    Detalle.imgPicture.Picture = LoadPicture(currentImg.path)
    Detalle.imgPicture.Stretch = True
    
    Detalle.Show
End Sub

Private Sub Form_Click()
    ' Desenfocar img
    If Not currentImg.img Is Nothing Then
        currentImg.img.BorderStyle = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If Me.currentImg.img Is Nothing Then Exit Sub

    Select Case KeyAscii
        Case 87  ' W key
            Me.currentImg.img.Top = Me.currentImg.img.Top - 100
            Me.currentImg.change_pos_dimentions = True
        Case 119  ' w key
            Me.currentImg.img.Top = Me.currentImg.img.Top - 100
            Me.currentImg.change_pos_dimentions = True
        Case 65 ' A key
            Me.currentImg.img.Left = Me.currentImg.img.Left - 100
            Me.currentImg.change_pos_dimentions = True
        Case 97 ' a key
            Me.currentImg.img.Left = Me.currentImg.img.Left - 100
            Me.currentImg.change_pos_dimentions = True
        Case 83 ' S key
            Me.currentImg.img.Top = Me.currentImg.img.Top + 100
            Me.currentImg.change_pos_dimentions = True
        Case 115 ' s key
            Me.currentImg.img.Top = Me.currentImg.img.Top + 100
            Me.currentImg.change_pos_dimentions = True
        Case 68 ' D key
            Me.currentImg.img.Left = Me.currentImg.img.Left + 100
            Me.currentImg.change_pos_dimentions = True
        Case 100 ' d key
            Me.currentImg.img.Left = Me.currentImg.img.Left + 100
            Me.currentImg.change_pos_dimentions = True
            
        ' Aumentar width
        Case 89 ' Y key
            Me.currentImg.img.Width = Me.currentImg.img.Width + 100
            Me.currentImg.change_pos_dimentions = True
        Case 121 ' y key
            Me.currentImg.img.Width = Me.currentImg.img.Width + 100
            Me.currentImg.change_pos_dimentions = True
            
        ' Disminuir width
        Case 85 ' U key
            Me.currentImg.img.Width = Me.currentImg.img.Width - 100
            Me.currentImg.change_pos_dimentions = True
        Case 117 ' u key
            Me.currentImg.img.Width = Me.currentImg.img.Width - 100
            Me.currentImg.change_pos_dimentions = True
        
        ' Aumentar height
        Case 72 ' H key
            Me.currentImg.img.Height = Me.currentImg.img.Height + 100
            Me.currentImg.change_pos_dimentions = True
        Case 104 ' h key
            Me.currentImg.img.Height = Me.currentImg.img.Height + 100
            Me.currentImg.change_pos_dimentions = True
            
        ' Disminuir height
        Case 74 ' J key
            Me.currentImg.img.Height = Me.currentImg.img.Height - 100
            Me.currentImg.change_pos_dimentions = True
        Case 106 ' j key
            Me.currentImg.img.Height = Me.currentImg.img.Height - 100
            Me.currentImg.change_pos_dimentions = True
    End Select
    
End Sub


Private Function loadPropositosFromDB()
    
    With gSqlConnection.miSql.SQLCommand
        .Clear
        .Add "SELECT * FROM propositos WITH(NOLOCK)"
    End With

    gSqlConnection.miSql.ExecSQL
    
    If gSqlConnection.miSql.HasRecords Then
    
        Do While Not gSqlConnection.miSql.EOF
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(0))   ' id
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(1))   ' date
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(2))    ' titulo
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(3))    ' descripcion
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(4))    ' categoria
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(5))    ' local_path (path from db)
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(6))    ' is_new
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(7))    ' img_width
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(8))    ' img_height
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(9))    ' img_top
            Debug.Print (gSqlConnection.miSql.Recordset.Fields(10))    ' img_left
            
            ' agregar imagen desde db
            AddImageFromDB gSqlConnection.miSql.Recordset.Fields(5), _
                            gSqlConnection.miSql.Recordset.Fields(2), _
                            gSqlConnection.miSql.Recordset.Fields(3), _
                            gSqlConnection.miSql.Recordset.Fields(4), _
                            gSqlConnection.miSql.Recordset.Fields(7), _
                            gSqlConnection.miSql.Recordset.Fields(8), _
                            gSqlConnection.miSql.Recordset.Fields(9), _
                            gSqlConnection.miSql.Recordset.Fields(10), _
                            gSqlConnection.miSql.Recordset.Fields(6), _
                            gSqlConnection.miSql.Recordset.Fields(1)
            
            gSqlConnection.miSql.MoveNext
        Loop
    
    End If
    
End Function


Private Sub Form_Load()
    imgsPath = "\imgs"
    Me.KeyPreview = True
    Me.countImages = 0
    Set Me.currentImg = New AgeusImg
    ReDim imgWrapper(0)
    Set Me.gSqlConnection = New DbWrapper
    
    ' Conectar a la base de datos
    If Me.gSqlConnection.ConnectDB Then
        Me.gSqlConnection.select100
    Else
        MsgBox "No se pudo conectar a la base de datos", vbCritical
        End
    End If
    
    loadCategorias
    loadPropositosFromDB
    
End Sub

Private Sub cmdAgregarActividad_Click()
    
    Dim frmAgregarAct As frmAgregarActividad
    Set frmAgregarAct = New frmAgregarActividad
    Set frmAgregarAct.parentFrm = Me
    frmAgregarAct.Show 1, Me
    
End Sub

' Se busca tanto localmente como en db si existe ya el proposito
Public Function existeTitle(title As String)

    ' Consulta sql
    With gSqlConnection.miSql.SQLCommand
        .Clear
        .Add "select top 1 1 from propositos with(nolock) where titulo like '" & title & "'"
    End With

    gSqlConnection.miSql.ExecSQL

    If gSqlConnection.miSql.HasRecords Then
        existeTitle = True
        Exit Function
    End If
    
    ' Consulta local
    If UBound(imgWrapper) > 0 Then
        Dim i As Integer
        For i = 1 To UBound(imgWrapper)
            If imgWrapper(i).title = title Then
                existeTitle = True
                Exit Function
            End If
        Next i
    End If
    
    existeTitle = False

End Function


Public Function AddImageFromDB(path As String, _
                               title As String, _
                               desc As String, _
                               categ As Integer, _
                               img_width As Integer, _
                               img_height As Integer, _
                               img_top As Integer, _
                               img_left As Integer, _
                               is_new As Byte, _
                               img_date As Date)

    ' obtener el path
    path = App.path & path
    
    ' obtener nombre de la imagen
    Dim imgName As String
    imgName = "img" & CStr(Me.countImages)

    ' dinamicamente agregar imagen al form
    Dim createdImg As Image
    Set createdImg = Me.Controls.Add("vb.Image", imgName, Me)
    createdImg.Width = img_width
    createdImg.Height = img_height
    createdImg.Stretch = True
    createdImg.ZOrder 1
    createdImg.Picture = LoadPicture(path)
    createdImg.Appearance = 0
    createdImg.Top = img_top
    createdImg.Left = img_left
    createdImg.Visible = True
    createdImg.ToolTipText = "Actividad: " & title & vbCrLf & vbCrLf & _
                                "Fecha: " & img_date & vbCrLf & vbCrLf & _
                                "Categoria: " & categ & vbCrLf & vbCrLf & _
                                "Descripcion: " & desc


    ' crear una instancia de mi wrapper de imagen y asignarle un id, imagen y el parent form
    Dim imgAgus As New AgeusImg
    Set imgAgus.img = createdImg
    Set imgAgus.parent = Me
    imgAgus.index = Me.countImages
    imgAgus.title = title
    imgAgus.descripcion = desc
    imgAgus.categoria = categ
    imgAgus.path = path
    imgAgus.isNew = is_new
    imgAgus.fecha = img_date
    
    ' setear como la current image a la imagen recien agregada
    Set Me.currentImg = imgAgus
    
    ' agrandar array de wrappers imagen
    ReDim Preserve imgWrapper(UBound(imgWrapper) + 1)
    
    ' agregar el wrapper actual
    Set imgWrapper(UBound(imgWrapper)) = imgAgus
    
    ' settear a nada el wrapper actual
    Set imgAgus = Nothing
    
    ' incrementar indeice
    Me.countImages = Me.countImages + 1

End Function


Public Function AddImage(path As String, _
                         title As String, _
                         desc As String, _
                         categ As Integer, _
                         img_date As Date)

    ' Si ya existe ese titulo de proposito retornar
    If existeTitle(title) Then
        MsgBox "Ya existe una actividad así", vbCritical
        Exit Function
    End If

    ' obtener nombre de la imagen
    Dim imgName As String
    imgName = "img" & CStr(Me.countImages)

    ' dinamicamente agregar imagen al form
    Dim createdImg As Image
    Set createdImg = Me.Controls.Add("vb.Image", imgName, Me)
    createdImg.Width = 3000
    createdImg.Height = 3000
    createdImg.Stretch = True
    createdImg.ZOrder 1
    createdImg.Picture = LoadPicture(path)
    createdImg.Appearance = 0
    createdImg.Top = 0
    createdImg.Left = 0
    createdImg.Visible = True
    createdImg.ToolTipText = "Actividad: " & title & vbCrLf & vbCrLf & _
                                "Fecha: " & img_date & vbCrLf & vbCrLf & _
                                "Categoria: " & categ & vbCrLf & vbCrLf & _
                                "Descripcion: " & desc
    
    ' crear una instancia de mi wrapper de imagen y asignarle un id, imagen y el parent form
    Dim imgAgus As New AgeusImg
    Set imgAgus.img = createdImg
    Set imgAgus.parent = Me
    imgAgus.index = Me.countImages
    imgAgus.title = title
    imgAgus.descripcion = desc
    imgAgus.categoria = categ
    imgAgus.path = path
    imgAgus.isNew = 1
    imgAgus.fecha = img_date
    
    ' setear como la current image a la imagen recien agregada
    Set Me.currentImg = imgAgus
    
    ' agrandar array de wrappers imagen
    ReDim Preserve imgWrapper(UBound(imgWrapper) + 1)
    
    ' agregar el wrapper actual
    Set imgWrapper(UBound(imgWrapper)) = imgAgus
    
    ' settear a nada el wrapper actual
    Set imgAgus = Nothing
    
    ' incrementar indeice
    Me.countImages = Me.countImages + 1

End Function
