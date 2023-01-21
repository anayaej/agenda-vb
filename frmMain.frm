VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LVTarjeta 
      Height          =   2415
      Left            =   90
      TabIndex        =   1
      Top             =   5220
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   176387
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4260
      Left            =   97
      TabIndex        =   0
      Top             =   900
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   7514
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Direccion"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Telefono (Hogar)"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Telefono (Celular)"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Telefono (Trabajo)"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "E-Mail"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Empesa"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Puesto"
         Object.Width           =   2822
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdEliminar 
      Height          =   735
      Left            =   2340
      TabIndex        =   2
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Eliminar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":058A
      PICN            =   "frmMain.frx":05A6
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
      Left            =   165
      TabIndex        =   3
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Agregar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":0940
      PICN            =   "frmMain.frx":095C
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
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Editar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":0CF6
      PICN            =   "frmMain.frx":0D12
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
      Left            =   3420
      TabIndex        =   5
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Buscar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":12AC
      PICN            =   "frmMain.frx":12C8
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
      Left            =   8910
      TabIndex        =   6
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":1862
      PICN            =   "frmMain.frx":187E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdEventos 
      Height          =   735
      Left            =   4500
      TabIndex        =   7
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Eventos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":1E18
      PICN            =   "frmMain.frx":1E34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdConfiguracion 
      Height          =   735
      Left            =   7830
      TabIndex        =   8
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Opciones"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmMain.frx":23CE
      PICN            =   "frmMain.frx":23EA
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
Dim DAMESODA As Integer

Private Sub ChameleonBtn1_Click()

frmContacto.Show

End Sub

Private Sub cmdConfiguracion_Click()

frmOpciones.Show vbModal, Me

End Sub

Private Sub cmdEventos_Click()

frmEventos.Show vbModal, Me

End Sub

Private Sub Command1_Click()
Dim sss As String
sss = Int(1 + 1.1)

MsgBox sss

End Sub

Private Sub Form_Initialize()

InitCommonControls
Me.Top = Screen.Height * 0.1
Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub cmdBuscar_Click()

frmBuscar.Show , Me

End Sub

Private Sub cmdEliminar_Click()
      
    If LV.ListItems.Count = 0 Then
        MsgBox "No existen registros en la base de datos", vbOKOnly + vbInformation, "Error"
        Exit Sub
    End If

    If MsgBox("Desea eliminar este registro?", vbYesNo + vbExclamation, "Eliminar registro") = vbYes Then
         ID = LV.SelectedItem.Text
         cnn.Execute "DELETE FROM Datos WHERE ID = " & ID & ""
         LV.ListItems.Remove (LV.SelectedItem.Index)
         LVTarjeta.ListItems.Clear
         Call RefrescarLista
         Rs.Close
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
    Dim fecha As String
    fecha = DateValue(Now)
        
    Call IniciarConexion
        
    Rs2.Open "SELECT * FROM Eventos WHERE Estado LIKE 'Activo' AND Fecha LIKE '" & fecha & "'", cnn
    If Rs2.RecordCount > 0 Then
        frmAlarma.Caption = "Ud. tiene" & " " & Rs2.RecordCount & " " & "evento(s) programados para hoy"
        DAMESODA = 1
    End If
    If Rs2.State = 1 Then Rs2.Close
        
    Rs.Open "SELECT * FROM Datos", cnn, adOpenDynamic, adLockOptimistic
    Call RefrescarLista
    Rs.Close
    
    Me.Show
    If DAMESODA = 1 Then frmAlarma.Show vbModal, Me
    
End Sub

Sub LlenarLVTarjeta()

Dim num_header As Integer
Dim ItemTarjeta As ListItem

    For num_header = 2 To LV.ColumnHeaders.Count
        Set ItemTarjeta = LVTarjeta.ListItems.Add(, , LV.ColumnHeaders(num_header))
        ItemTarjeta.SubItems(1) = LV.SelectedItem.SubItems(num_header - 1)
    Next

End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True

    If LV.SortOrder = lvwAscending Then
        LV.SortOrder = lvwDescending
        ElseIf LV.SortOrder = lvwDescending Then
        LV.SortOrder = lvwAscending
    End If

End Sub

Private Sub LV_ItemClick(ByVal item As MSComctlLib.ListItem)

    LVTarjeta.ListItems.Clear
    Call LlenarLVTarjeta

End Sub

Sub Agregar()

frmDatos.Caption = "Agregar Entrada"
Modo = "Agregar"
frmDatos.Show vbModal

End Sub
Sub Editar()
Dim i As Integer

    If frmMain.LV.ListItems.Count = 0 Then
        MsgBox "No existen registros en la base de datos", vbOKOnly + vbInformation, "Error"
        Exit Sub
    End If
    For i = 1 To 9
        frmDatos.txtDatos(i - 1) = frmMain.LV.SelectedItem.SubItems(i)
    Next
    frmDatos.txtId = frmMain.LV.SelectedItem
    frmDatos.Caption = "Editar Entrada"
    Modo = "Editar"
    frmDatos.Show vbModal
End Sub
