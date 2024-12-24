VERSION 5.00
Begin VB.Form Detalle 
   Caption         =   "Detalle"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
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
   ScaleHeight     =   7365
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFecha 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtDesc 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5880
      Width           =   6615
   End
   Begin VB.TextBox txtCategoria 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtTitulo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Image imgPicture 
      Height          =   4455
      Left            =   480
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblPropositoFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposito fecha:"
      Height          =   195
      Left            =   4080
      TabIndex        =   3
      Top             =   5160
      Width           =   1620
   End
   Begin VB.Label lblPropositoDescripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposito descripcion:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   5520
      Width           =   1620
   End
   Begin VB.Label lblPropositoCategoria 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposito categoria:"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label lblPropositoTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposito titulo:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   4800
      Width           =   1140
   End
End
Attribute VB_Name = "Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
