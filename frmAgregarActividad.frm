VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgregarActividad 
   Caption         =   "Agregando Actividad"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
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
   ScaleHeight     =   6060
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdBuscarImagen 
      Left            =   720
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAgregarActividad 
      Caption         =   "Agregar Actividad"
      Height          =   480
      Left            =   2040
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscarImagen 
      Caption         =   "Buscar Imagen"
      Height          =   480
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin MSComCtl2.MonthView mvActividad 
      Height          =   2370
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   177078273
      CurrentDate     =   45637
   End
   Begin VB.ComboBox cboCategoria 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtDescripcionAct 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtTituloActividad 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label lblFechaEstimada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha estimada de realizaciòn"
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   720
      Width           =   2145
   End
   Begin VB.Label lblCategoría 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label lblDescripcionDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion de la actividad"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1890
   End
   Begin VB.Label lblTituloDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo de la actividad"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "frmAgregarActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public parentFrm As VisionBoard
Public imgPath As String

Private Sub cmdAgregarActividad_Click()
    ' Validar campos
    If Me.txtTituloActividad = "" Then
        MsgBox "Agrega un título a la actividad"
        Exit Sub
    ElseIf Me.txtDescripcionAct = "" Then
        MsgBox "Agrega una descripcion a la actividad"
        Exit Sub
    ElseIf Me.imgPath = "" Then
        MsgBox "Agrega una imagen"
        Exit Sub
    End If
    
    parentFrm.AddImage imgPath, txtTituloActividad, txtDescripcionAct, cboCategoria.ListIndex, mvActividad
    MsgBox "Actividad agregada con exito, puede ir a posicionarla redimencionarla"
    Unload Me
    
End Sub

Private Sub cmdBuscarImagen_Click()
    ' Obtenemos el path de la imagen, solo acepta jpg
    With cdBuscarImagen
        .DialogTitle = "Cargar Imagen :D"
        .InitDir = "C:\"
        .Filter = "Text Files (*.jpg)|*.jpg"
        .ShowOpen
    End With
    
    Me.imgPreview.Picture = LoadPicture(cdBuscarImagen.FileName)
    Me.imgPath = cdBuscarImagen.FileName
End Sub

Public Sub Form_Load()

    ' Cargar combobox
    With parentFrm.gSqlConnection.miSql.SQLCommand
        .Clear
        .Add "select * from categorias_proposito with(nolock)"
    End With
    
    parentFrm.gSqlConnection.miSql.ExecSQL
    If parentFrm.gSqlConnection.miSql.HasRecords Then
        Do While Not parentFrm.gSqlConnection.miSql.EOF
            cboCategoria.AddItem parentFrm.gSqlConnection.miSql.Recordset.Fields(1), parentFrm.gSqlConnection.miSql.Recordset.Fields(0)
            parentFrm.gSqlConnection.miSql.MoveNext
        Loop
    End If

End Sub



