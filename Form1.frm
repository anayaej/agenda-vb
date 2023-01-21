VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonButton.ChameleonBtn cmdEliminar 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Eliminar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":058A
      PICN            =   "Form1.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionDatos 
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Agregar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":0940
      PICN            =   "Form1.frx":095C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5655
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Telefono"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Direccion"
         Object.Width           =   5292
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionDatos 
      Height          =   735
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Editar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":0CF6
      PICN            =   "Form1.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdBuscar 
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Buscar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":12AC
      PICN            =   "Form1.frx":12C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdSalir 
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":1862
      PICN            =   "Form1.frx":187E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Form_Initialize()
InitCommonControls
End Sub


Private Sub cmdBuscar_Click()

frmBuscar.Show , Me

End Sub

Private Sub cmdEliminar_Click()
      
    If LV.ListItems.Count = 0 Then
        MsgBox "No existen registros en la base de datos", vbOKOnly + vbInformation, "Error"
        Exit Sub
    End If

    If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
         "Nombre:" & LV.SelectedItem.SubItems(1) & vbNewLine & _
         "Apellido:" & LV.SelectedItem.SubItems(2) & vbNewLine & _
         "Telefono:" & LV.SelectedItem.SubItems(3) & vbNewLine & _
         "Direccion:" & LV.SelectedItem.SubItems(4), vbYesNo + vbExclamation, "Eliminacion de registro") = vbYes Then
         ID = LV.SelectedItem.Text
         cnn.Execute "DELETE FROM Datos WHERE ID = " & ID & ""
         LV.ListItems.Remove (LV.SelectedItem.Index)
         Call RefrescarLista
         
    End If
End Sub

Private Sub cmdOpcionDatos_Click(Index As Integer)
    Select Case Index
        Case 1: Call Agregar
        Case 2: Call Editar
    End Select
End Sub

Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()

    Call IniciarConexion
    Rs.Open "SELECT * FROM Datos", cnn, adOpenDynamic, adLockOptimistic
    Call RefrescarLista
    
End Sub



Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

LV.Sorted = True

If LV.SortOrder = lvwAscending Then
   LV.SortOrder = lvwDescending
   ElseIf LV.SortOrder = lvwDescending Then
   LV.SortOrder = lvwAscending
End If

End Sub
