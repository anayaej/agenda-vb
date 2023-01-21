VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEdEventos 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtId 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkAlarma 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Activar"
      Height          =   285
      Left            =   165
      TabIndex        =   8
      Top             =   3960
      Width           =   1635
   End
   Begin VB.TextBox txtNotas 
      Height          =   1725
      Left            =   165
      TabIndex        =   6
      Top             =   2070
      Width           =   4515
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1260
      Width           =   3075
   End
   Begin VB.ComboBox cboNombres 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   3075
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   16711681
      UpDown          =   -1  'True
      CurrentDate     =   40234
      MaxDate         =   73415
      MinDate         =   36526
   End
   Begin ChamaleonButton.ChameleonBtn cmdGuardar 
      Height          =   465
      Left            =   960
      TabIndex        =   9
      Top             =   4410
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Guardar"
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
      MICON           =   "frmEdEventos.frx":0000
      PICN            =   "frmEdEventos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   465
      Left            =   2580
      TabIndex        =   10
      Top             =   4410
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "Cancelar"
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
      MICON           =   "frmEdEventos.frx":05B6
      PICN            =   "frmEdEventos.frx":05D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "ID:"
      Height          =   375
      Index           =   4
      Left            =   3195
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Notas adicionales"
      Height          =   285
      Left            =   255
      TabIndex        =   7
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Tarea:"
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Contacto:"
      Height          =   285
      Left            =   300
      TabIndex        =   4
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EEEEDF&
      Caption         =   "Fecha:"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   180
      Width           =   645
   End
End
Attribute VB_Name = "frmEdEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub cmdGuardar_Click()
Dim Estado As String

If chkAlarma.Value = 1 Then Estado = "Activo"
If chkAlarma.Value = 0 Then Estado = "Inactivo"

    Select Case evModo
            
            Case "AGREGAR"
            cnn.Execute "INSERT INTO Eventos(Fecha,Nombre,Descripcion,NotasAdicionales,Estado) VALUES ('" & dtFecha.Value & "', '" & cboNombres.Text & "', '" & txtDescripcion.Text & "', '" & txtNotas.Text & "', '" & Estado & "')"
            Rs2.Open
            frmEventos.LVEventos.ListItems.Clear
            Call RefrescarListaEventos
            Rs2.Close
            evModo = ""
            
            Case "EDITAR"
            cnn.Execute "UPDATE Eventos SET Fecha = '" & dtFecha.Value & "', Nombre = '" & cboNombres.Text & "', Descripcion = '" & txtDescripcion.Text & "', NotasAdicionales = '" & txtNotas.Text & "', Estado = '" & Estado & "' WHERE ID = " & txtId.Text & ""
            Rs2.Open
            frmEventos.LVEventos.ListItems.Clear
            Call RefrescarListaEventos
            Rs2.Close
            evModo = ""

    End Select
Unload Me
End Sub

Private Sub txtDescripcion_Change()


    If BuscarCaracter(txtDescripcion.Text) = 1 Then
        txtDescripcion.Text = Left(txtDescripcion.Text, Len(txtDescripcion.Text) - 1)
        txtDescripcion.SelStart = Len(txtDescripcion.Text)
        Exit Sub
    End If

End Sub

Private Sub txtNotas_Change()

    If BuscarCaracter(txtNotas.Text) = 1 Then
        txtNotas.Text = Left(txtNotas.Text, Len(txtNotas.Text) - 1)
        txtNotas.SelStart = Len(txtNotas.Text)
        Exit Sub
    End If
End Sub
