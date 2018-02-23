VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmKeyboard 
   BorderStyle     =   4  '°íÁ¤ µµ±¸ Ã¢
   Caption         =   "KEYBOARD"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15870
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
   ScaleHeight     =   9525
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.TextBox txtInput 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      IMEMode         =   8  '¿µ¹®
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   14295
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   0
      Left            =   12060
      TabIndex        =   0
      Tag             =   "0"
      Top             =   5880
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1931
      Caption         =   "0"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   1
      Left            =   12060
      TabIndex        =   2
      Tag             =   "1"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "1"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   2
      Left            =   13200
      TabIndex        =   3
      Tag             =   "2"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "2"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   3
      Left            =   14340
      TabIndex        =   4
      Tag             =   "3"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "3"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   4
      Left            =   12060
      TabIndex        =   5
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "4"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   5
      Left            =   13200
      TabIndex        =   6
      Tag             =   "5"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "5"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   6
      Left            =   14340
      TabIndex        =   7
      Tag             =   "6"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "6"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   7
      Left            =   12060
      TabIndex        =   8
      Tag             =   "7"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "7"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   8
      Left            =   13200
      TabIndex        =   9
      Tag             =   "8"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "8"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   9
      Left            =   14340
      TabIndex        =   10
      Tag             =   "9"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "9"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnNumKey 
      Height          =   1095
      Index           =   10
      Left            =   14340
      TabIndex        =   11
      Tag             =   "-"
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "-"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Tag             =   "Q"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "Q"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   1
      Left            =   1620
      TabIndex        =   13
      Tag             =   "W"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "W"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Tag             =   "E"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "E"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   3
      Left            =   3900
      TabIndex        =   15
      Tag             =   "R"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "R"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   4
      Left            =   5040
      TabIndex        =   16
      Tag             =   "T"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "T"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   5
      Left            =   6180
      TabIndex        =   17
      Tag             =   "Y"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "Y"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   6
      Left            =   7320
      TabIndex        =   18
      Tag             =   "U"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "U"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   7
      Left            =   8460
      TabIndex        =   19
      Tag             =   "I"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "I"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   8
      Left            =   9600
      TabIndex        =   20
      Tag             =   "O"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "O"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   9
      Left            =   10740
      TabIndex        =   21
      Tag             =   "P"
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "P"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   10
      Left            =   1020
      TabIndex        =   22
      Tag             =   "A"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "A"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   11
      Left            =   2160
      TabIndex        =   23
      Tag             =   "S"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "S"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   12
      Left            =   3300
      TabIndex        =   24
      Tag             =   "D"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "D"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   13
      Left            =   4440
      TabIndex        =   25
      Tag             =   "F"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "F"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   14
      Left            =   5580
      TabIndex        =   26
      Tag             =   "G"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "G"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   15
      Left            =   6720
      TabIndex        =   27
      Tag             =   "H"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "H"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   16
      Left            =   7860
      TabIndex        =   28
      Tag             =   "J"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "J"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   17
      Left            =   9000
      TabIndex        =   29
      Tag             =   "K"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "K"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   18
      Left            =   10140
      TabIndex        =   30
      Tag             =   "L"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "L"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   19
      Left            =   1620
      TabIndex        =   31
      Tag             =   "Z"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "Z"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   20
      Left            =   2760
      TabIndex        =   32
      Tag             =   "X"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "X"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   21
      Left            =   3900
      TabIndex        =   33
      Tag             =   "C"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "C"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   22
      Left            =   5040
      TabIndex        =   34
      Tag             =   "V"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "V"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   23
      Left            =   6180
      TabIndex        =   35
      Tag             =   "B"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "B"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   24
      Left            =   7320
      TabIndex        =   36
      Tag             =   "N"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "N"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   25
      Left            =   8460
      TabIndex        =   37
      Tag             =   "M"
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "M"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   29
      Left            =   9600
      TabIndex        =   38
      Tag             =   "M"
      Top             =   5880
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1931
      Caption         =   "DEL"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1575
      Index           =   30
      Left            =   480
      TabIndex        =   39
      Tag             =   "M"
      Top             =   7500
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   2778
      Caption         =   "ENTER"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   26
      Left            =   480
      TabIndex        =   40
      Tag             =   "M"
      Top             =   5880
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1931
      Caption         =   "SHIFT"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   31
      Left            =   2760
      TabIndex        =   41
      Tag             =   "M"
      Top             =   5880
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1931
      Caption         =   "SPACE"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton btnEngKey 
      Height          =   1095
      Index           =   33
      Left            =   10740
      TabIndex        =   42
      Tag             =   "."
      Top             =   4740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "."
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      BackColor       =   15790320
      ImgOutLineSize  =   3
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000005&
      BorderStyle     =   1  '´ÜÀÏ °íÁ¤
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   480
      TabIndex        =   43
      Top             =   600
      Width           =   14955
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    txtInput.PasswordChar = KeyBoardObj.PasswordChar
    KeyBoardObj.Tag = KeyBoardObj.Text
    txtInput.Text = KeyBoardObj.Text

'    txtInput.IMEMode = 8
'    btnEngKey(28).Caption = "Kor" '"ÇÑ±Û"
    
    If bKeyBoardNum = True Then
        btnEngKey(33).Enabled = True
    Else
        btnEngKey(33).Enabled = False
    End If
    
    txtInput.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then
        KeyBoardObj.Text = KeyBoardObj.Tag
    End If
End Sub

Private Sub btnNumKey_Click(Index As Integer)
    Call AddTouchChar(Index, btnNumKey(Index).Tag)
End Sub

Private Sub btnEngKey_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 26:        ' caps lock
            If btnEngKey(0).Caption = "Q" Then
                For i = 0 To 25
                    btnEngKey(i).Caption = LCase(btnEngKey(i).Caption)
                Next
            Else
                For i = 0 To 25
                    btnEngKey(i).Caption = UCase(btnEngKey(i).Caption)
                Next
            End If
                    
        Case 28:    ' han / eng
'            If txtInput.IMEMode <> 9 Then
'                txtInput.IMEMode = 9
'                cmdEngKey(28).Caption = "¿µ¾î"
'            Else
'                txtInput.IMEMode = 8
'                cmdEngKey(28).Caption = "ÇÑ±Û"
'            End If
                    
        Case 30:    ' Enter
            If bKeyBoardNum = True Then
                If IsNumeric(txtInput.Text) = True Then
                    KeyBoardObj.Tag = txtInput.Text
                End If
            Else
                KeyBoardObj.Tag = txtInput.Text
            End If
            
            Debug.Print "INPUT : " & txtInput.Text
            Unload Me
                    
        Case 31: Call AddTouchChar(Index, " ")  ' Space
        
        Case 29: Call VirtualKeyDelete          ' Delete
                    
        Case Else:
            If btnEngKey(0).Caption = "Q" Then
                Call AddTouchChar(Index, btnEngKey(Index).Tag)
            Else
                Call AddTouchChar(Index, LCase(btnEngKey(Index).Tag))
            End If
    End Select
End Sub

' Function List ===============================================================

Private Sub AddTouchChar(ByVal Index As Integer, ByVal lpChar As String)
    txtInput.Text = txtInput.Text & lpChar
End Sub

Private Sub VirtualKeyDelete()
    Dim lpInput As String
    Dim nLen    As Integer
    
    lpInput = txtInput.Text
    nLen = Len(txtInput.Text)
    
    If (nLen > 0) Then
        txtInput.Text = Left(lpInput, nLen - 1)
    End If
End Sub

