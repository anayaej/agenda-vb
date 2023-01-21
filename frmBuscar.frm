VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmBuscar 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5235
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   405
      Left            =   3862
      TabIndex        =   3
      Top             =   300
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Cerrar"
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
      MICON           =   "frmBuscar.frx":628A
      PICN            =   "frmBuscar.frx":62A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optBuscarEn 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Nombre"
      Height          =   225
      Index           =   0
      Left            =   1282
      TabIndex        =   1
      Top             =   660
      Width           =   1065
   End
   Begin VB.OptionButton optBuscarEn 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Apellido"
      Height          =   225
      Index           =   1
      Left            =   2632
      TabIndex        =   2
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtBuscar 
      Height          =   315
      Left            =   172
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Buscar en:"
      Height          =   195
      Left            =   202
      TabIndex        =   4
      Top             =   660
      Width           =   1005
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    optBuscarEn(0) = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Rs.Open "SELECT * FROM DATOS", cnn, adOpenDynamic, adLockOptimistic
    Call RefrescarLista
    Rs.Close
    Unload Me

End Sub

Private Sub txtBuscar_Change()

    If BuscarCaracter(txtBuscar.Text) = 1 Then
        txtBuscar.Text = Left(txtBuscar.Text, Len(txtBuscar.Text) - 1)
        txtBuscar.SelStart = Len(txtBuscar.Text)
        Exit Sub
    Else
        Select Case True
            Case optBuscarEn(0).Value
                If Rs.State = 1 Then Rs.Close
                Rs.Open "SELECT * FROM Datos WHERE Nombre LIKE '%" & txtBuscar.Text & "%'"
                Call RefrescarLista
                Rs.Close
            Case optBuscarEn(1).Value
                If Rs.State = 1 Then Rs.Close
                Rs.Open "SELECT * FROM Datos WHERE Apellido LIKE '%" & txtBuscar.Text & "%'"
                Call RefrescarLista
                Rs.Close
        End Select
    End If


End Sub

