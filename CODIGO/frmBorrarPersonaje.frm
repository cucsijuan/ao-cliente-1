VERSION 5.00
Begin VB.Form frmBorrarPersonaje 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrar personaje"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtPasswd 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "Clave:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNombre 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "Nick:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   3000
      Tag             =   "1"
      Top             =   1800
      Width           =   930
   End
End
Attribute VB_Name = "frmBorrarPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
        UserName = txtNombre.Text
        UserEmail = txtMail.Text
        UserClave = txtClave.Text
        UserPassword = txtPasswd.Text
       
        If Not CheckMailString(txtMail.Text) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
        End If
       
        If UserName = vbNullString Then
            MsgBox "Ingrese el nick de su personaje."
            Exit Sub
        End If
       
        If UserClave = vbNullString Then
            MsgBox "Ingrese la clave pin de su personaje."
            Exit Sub
        End If
       
        If UserPassword = vbNullString Then
            MsgBox "Ingrese la contraseña de su personaje."
            Exit Sub
        End If
        Call Login
        Unload Me
End Sub

