VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tipo de Pelicula"
      Top             =   3960
      Width           =   4095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   4680
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   " "
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2160
      TabIndex        =   11
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "alquiler"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "actor"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "pelicula"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "cassette"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "tipo_pelicula"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "cliente"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "alquiler"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "actor"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "pelicula"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "cassette"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "tipo_pelicula"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
Data1.Recordset.Update

End Sub

Private Sub Command3_Click()
Data1.Recordset.Delete

End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveLast

End Sub




