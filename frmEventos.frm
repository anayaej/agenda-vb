VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEventos 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   Icon            =   "frmEventos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   735
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Nuevo"
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
      MICON           =   "frmEventos.frx":058A
      PICN            =   "frmEventos.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LVEventos 
      Height          =   4695
      Left            =   90
      TabIndex        =   1
      Top             =   900
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   8281
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dias Restantes"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contacto"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tarea"
         Object.Width           =   4851
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Notas Adicionales"
         Object.Width           =   4851
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Estado"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtEventos 
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16777217
      CurrentDate     =   40234
      MaxDate         =   41274
      MinDate         =   40179
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   735
      Index           =   1
      Left            =   8550
      TabIndex        =   3
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmEventos.frx":0B40
      PICN            =   "frmEventos.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   735
      Index           =   2
      Left            =   9540
      TabIndex        =   4
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmEventos.frx":10F6
      PICN            =   "frmEventos.frx":1112
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
Attribute VB_Name = "frmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha As String

Private Sub cmdOpciones_Click(Index As Integer)

    Select Case Index
        Case 0:
            evModo = "AGREGAR"
            frmEdEventos.Caption = "Agregar evento"
            frmEdEventos.dtFecha.Value = Date
            Rs.Open "SELECT * FROM Datos", cnn
            If Rs.RecordCount > 0 Then
                Do While Rs.EOF = False
                frmEdEventos.cboNombres.AddItem Rs("Nombre") & " " & Rs("Apellido")
                Rs.MoveNext
                Loop
            End If
                       
            
            frmEdEventos.Show vbModal, Me
            
            Rs.Close
        Case 1:
            evModo = "EDITAR"
            
            If LVEventos.ListItems.Count = 0 Then
                MsgBox "No existen eventos para editar", vbInformation + vbOKOnly, "Error"
                Exit Sub
            End If
            
            frmEdEventos.txtId.Text = LVEventos.SelectedItem
            frmEdEventos.dtFecha = LVEventos.SelectedItem.SubItems(1)
            frmEdEventos.cboNombres.Text = LVEventos.SelectedItem.SubItems(3)
            frmEdEventos.txtDescripcion.Text = LVEventos.SelectedItem.SubItems(4)
            frmEdEventos.txtNotas.Text = LVEventos.SelectedItem.SubItems(5)
            
            If LVEventos.SelectedItem.SubItems(6) = "Activo" Then frmEdEventos.chkAlarma.Value = 1
            If LVEventos.SelectedItem.SubItems(6) = "Inactivo" Then frmEdEventos.chkAlarma.Value = 0
            
            Rs.Open "SELECT * FROM Datos", cnn
            If Rs.RecordCount > 0 Then
                Do While Rs.EOF = False
                frmEdEventos.cboNombres.AddItem Rs("Nombre") & " " & Rs("Apellido")
                Rs.MoveNext
                Loop
            End If
            Rs.Close
            frmEdEventos.Caption = "Editar evento"
            frmEdEventos.Show vbModal, Me
            
        Case 2:
            
            If LVEventos.ListItems.Count = 0 Then
                MsgBox "No existen eventos para eliminar", vbInformation + vbOKOnly, "Error"
                Exit Sub
            End If
            
            If MsgBox("Desea eliminar este evento?", vbOKCancel + vbInformation, "Eliminar evento") = vbOK Then
            cnn.Execute "DELETE FROM Eventos WHERE ID = " & LVEventos.SelectedItem & ""
            LVEventos.ListItems.Remove (LVEventos.SelectedItem.Index)
            End If
            
            
    End Select
End Sub



Private Sub dtEventos_Change(Index As Integer)

    LVEventos.ListItems.Clear
    fecha = dtEventos(0).Value
    Rs2.Open "SELECT * FROM Eventos WHERE Fecha =  # " + fecha + " #", cnn
        If Rs2.RecordCount > 0 Then
            Rs2.MoveFirst
           Call RefrescarListaEventos
    
        End If
        Rs2.Close
End Sub

Private Sub Form_Load()
Dim evento As ListItem

    dtEventos(0).Value = Date
    Rs2.Open "SELECT * FROM Eventos", cnn
    Call RefrescarListaEventos
    Rs2.Close

End Sub

Private Sub Label1_Click()

End Sub

