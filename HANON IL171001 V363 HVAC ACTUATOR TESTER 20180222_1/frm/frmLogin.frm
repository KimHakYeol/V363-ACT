VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '°íÁ¤ µµ±¸ Ã¢
   Caption         =   "LOGIN"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "³ª´®°íµñ"
      Size            =   12
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6390
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.CommandButton btnCancel 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00F0F0F0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   36
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   780
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   3
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton btnLogin 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00F0F0F0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   48
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   3960
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   2
      Top             =   3720
      Width           =   5955
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   48
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      IMEMode         =   3  '»ç¿ë ¸øÇÔ
      Left            =   780
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2220
      Width           =   9135
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   48
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   840
      TabIndex        =   0
      Top             =   900
      Width           =   5355
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bLoading As Boolean

Private Sub Form_Activate()
    nNowForm = FM_LOGIN
    
    Call LoadLangFile(FM_LOGIN)
    
    If bLoading = False Then
        bLoading = True
        
        bLogin = False
        txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Load()
    bLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub

Private Sub btnLogin_Click()
    Call ConfirmPassword
End Sub

Private Sub btnCancel_Click()
    bLogin = False      ' Always Question
    Unload Me
End Sub

Private Sub txtPassword_Click()
    Call CheckKeyBoardDisplay(txtPassword, False)
End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call ConfirmPassword
    End If
End Sub

' Function List ===============================================================

Private Sub ConfirmPassword()
    Dim lpValue As String
    
    lpValue = Trim$(txtPassword.Text)
    
    If lpValue = Trim$(SysVar.lpPassword) Or UCase(Trim$(lpValue)) = MASTER_PASSWORD Then
        bLogin = True
        Unload Me
    Else
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
End Sub

