VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmDatos 
   BackColor       =   &H00EEEEDF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5400
   Icon            =   "frmDatos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4652.875
   ScaleMode       =   0  'User
   ScaleWidth      =   5028.91
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   6
      Left            =   1545
      TabIndex        =   15
      Top             =   3270
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   8
      Left            =   1545
      TabIndex        =   17
      Top             =   4230
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   7
      Left            =   1545
      TabIndex        =   16
      Top             =   3750
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   5
      Left            =   1545
      TabIndex        =   14
      Top             =   2790
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   4
      Left            =   1545
      TabIndex        =   13
      Top             =   2310
      Width           =   3375
   End
   Begin ChamaleonButton.ChameleonBtn cmdGuardar 
      Height          =   465
      Left            =   1635
      TabIndex        =   18
      Top             =   4770
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
      MICON           =   "frmDatos.frx":058A
      PICN            =   "frmDatos.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
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
      Left            =   825
      TabIndex        =   20
      Top             =   4860
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   0
      Left            =   1545
      TabIndex        =   9
      Top             =   330
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   1
      Left            =   1545
      TabIndex        =   10
      Top             =   810
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   3
      Left            =   1545
      TabIndex        =   12
      Top             =   1770
      Width           =   3375
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   1260
      Width           =   3375
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   465
      Left            =   3255
      TabIndex        =   19
      Top             =   4770
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
      MICON           =   "frmDatos.frx":0B40
      PICN            =   "frmDatos.frx":0B5C
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
      Caption         =   "Correo Electronico:"
      Height          =   375
      Index           =   10
      Left            =   225
      TabIndex        =   6
      Top             =   3150
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Puesto:"
      Height          =   375
      Index           =   8
      Left            =   465
      TabIndex        =   8
      Top             =   4230
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Empresa:"
      Height          =   375
      Index           =   7
      Left            =   345
      TabIndex        =   7
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Telefono de trabajo:"
      Height          =   375
      Index           =   6
      Left            =   225
      TabIndex        =   5
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Telefono celular:"
      Height          =   375
      Index           =   5
      Left            =   285
      TabIndex        =   4
      Top             =   2190
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   285
      TabIndex        =   21
      Top             =   4860
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Nombre:"
      Height          =   375
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   270
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Apellido:"
      Height          =   375
      Index           =   1
      Left            =   465
      TabIndex        =   1
      Top             =   870
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Telefono:"
      Height          =   375
      Index           =   2
      Left            =   345
      TabIndex        =   3
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDF&
      Caption         =   "Direccion:"
      Height          =   375
      Index           =   3
      Left            =   330
      TabIndex        =   2
      Top             =   1350
      Width           =   975
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub cmdGuardar_Click()
Dim i As Integer
    Select Case Modo
        Case "Editar"
            For i = 0 To txtDatos.UBound
            If BuscarCaracter(txtDatos(i)) = 1 Then
                Exit Sub
            End If
            Next
            cnn.Execute "UPDATE Datos SET Nombre = '" & txtDatos(0).Text & "', Apellido = '" & txtDatos(1).Text & "', Direccion = '" & txtDatos(2).Text & "', Thogar = '" & txtDatos(3).Text & "', TCelular = '" & txtDatos(4).Text & "',TTrabajo = '" & txtDatos(5).Text & "',  Email = '" & txtDatos(6).Text & "', Empresa = '" & txtDatos(7).Text & "', Puesto = '" & txtDatos(8).Text & "'  WHERE ID = " & txtId.Text & ""
            Call RefrescarLista
            Rs.Close
            Unload Me
        
        Case "Agregar"
                        
            For i = 0 To txtDatos.UBound
            If BuscarCaracter(txtDatos(i)) = 1 Then
                Exit Sub
            End If
            Next
            cnn.Execute "INSERT INTO Datos(Nombre, Apellido, Direccion, THogar, TCelular, TTrabajo, Email, Empresa, Puesto) VALUES ('" & txtDatos(0).Text & "', '" & txtDatos(1) & "', '" & txtDatos(2) & "', '" & txtDatos(3) & "', '" & txtDatos(4) & "', '" & txtDatos(5) & "', '" & txtDatos(6) & "', '" & txtDatos(7) & "', '" & txtDatos(8) & "')"
            Call RefrescarLista
            Rs.Close
            Unload Me
            
    End Select

End Sub





