VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgCancelar 
      Height          =   285
      Left            =   960
      TabIndex        =   22
      Top             =   6000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgAceptar 
      Height          =   285
      Left            =   5040
      TabIndex        =   23
      Top             =   6000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   24
      Top             =   840
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   25
      Top             =   1200
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   26
      Top             =   1560
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   27
      Top             =   1920
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   28
      Top             =   2280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   29
      Top             =   2640
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   30
      Top             =   3000
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   31
      Top             =   3360
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   32
      Top             =   3720
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   33
      Top             =   4080
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   11
      Left            =   7320
      TabIndex        =   34
      Top             =   840
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   12
      Left            =   7320
      TabIndex        =   35
      Top             =   1200
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   13
      Left            =   7320
      TabIndex        =   36
      Top             =   1560
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   14
      Left            =   7320
      TabIndex        =   37
      Top             =   1920
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   15
      Left            =   7320
      TabIndex        =   38
      Top             =   2280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   16
      Left            =   7320
      TabIndex        =   39
      Top             =   2640
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   17
      Left            =   7320
      TabIndex        =   40
      Top             =   3000
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   18
      Left            =   7320
      TabIndex        =   41
      Top             =   3360
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   19
      Left            =   7320
      TabIndex        =   42
      Top             =   3720
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMenos 
      Height          =   285
      Index           =   20
      Left            =   7320
      TabIndex        =   43
      Top             =   4080
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   44
      Top             =   840
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   45
      Top             =   1200
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   46
      Top             =   1560
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   47
      Top             =   1920
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   48
      Top             =   2280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   49
      Top             =   2640
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   50
      Top             =   3000
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   51
      Top             =   3360
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   52
      Top             =   3720
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   53
      Top             =   4080
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   11
      Left            =   8160
      TabIndex        =   54
      Top             =   840
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   12
      Left            =   8160
      TabIndex        =   55
      Top             =   1200
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   13
      Left            =   8160
      TabIndex        =   56
      Top             =   1560
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   14
      Left            =   8160
      TabIndex        =   57
      Top             =   1920
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   15
      Left            =   8160
      TabIndex        =   58
      Top             =   2280
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   16
      Left            =   8160
      TabIndex        =   59
      Top             =   2640
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   17
      Left            =   8160
      TabIndex        =   60
      Top             =   3000
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   18
      Left            =   8160
      TabIndex        =   61
      Top             =   3360
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   19
      Left            =   8160
      TabIndex        =   62
      Top             =   3720
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgMas 
      Height          =   285
      Index           =   20
      Left            =   8160
      TabIndex        =   63
      Top             =   4080
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin VB.Image imgNavegacion 
      Height          =   375
      Left            =   4695
      Top             =   4110
      Width           =   1440
   End
   Begin VB.Image imgCombateSinArmas 
      Height          =   345
      Left            =   4695
      Top             =   3735
      Width           =   2100
   End
   Begin VB.Image imgCombateDistancia 
      Height          =   345
      Left            =   4695
      Top             =   3345
      Width           =   2280
   End
   Begin VB.Image imgDomar 
      Height          =   345
      Left            =   4695
      Top             =   2970
      Width           =   1845
   End
   Begin VB.Image imgLiderazgo 
      Height          =   330
      Left            =   4695
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Image imgHerreria 
      Height          =   345
      Left            =   4695
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Image imgCarpinteria 
      Height          =   360
      Left            =   4695
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Image imgMineria 
      Height          =   360
      Left            =   4695
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Image imgPesca 
      Height          =   330
      Left            =   4695
      Top             =   1110
      Width           =   780
   End
   Begin VB.Image imgEscudos 
      Height          =   270
      Left            =   4695
      Top             =   720
      Width           =   2340
   End
   Begin VB.Image imgComercio 
      Height          =   330
      Left            =   495
      Top             =   4125
      Width           =   1170
   End
   Begin VB.Image imgTalar 
      Height          =   360
      Left            =   495
      Top             =   3750
      Width           =   885
   End
   Begin VB.Image imgSupervivencia 
      Height          =   330
      Left            =   495
      Top             =   3375
      Width           =   1620
   End
   Begin VB.Image imgOcultarse 
      Height          =   345
      Left            =   495
      Top             =   3030
      Width           =   1230
   End
   Begin VB.Image imgApunialar 
      Height          =   360
      Left            =   495
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Image imgMeditar 
      Height          =   345
      Left            =   495
      Top             =   2265
      Width           =   1065
   End
   Begin VB.Image imgCombateArmas 
      Height          =   315
      Left            =   495
      Top             =   1890
      Width           =   2280
   End
   Begin VB.Image imgEvasion 
      Height          =   330
      Left            =   495
      Top             =   1515
      Width           =   2295
   End
   Begin VB.Image imgRobar 
      Height          =   360
      Left            =   495
      Top             =   1125
      Width           =   930
   End
   Begin VB.Image imgMagia 
      Height          =   330
      Left            =   495
      Top             =   750
      Width           =   870
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   600
      TabIndex        =   21
      Top             =   4800
      Width           =   7815
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   20
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   19
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   18
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   17
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   15
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   14
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   13
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   12
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   11
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   10
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7680
      TabIndex        =   9
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7680
      TabIndex        =   8
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7680
      TabIndex        =   7
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7680
      TabIndex        =   6
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7680
      TabIndex        =   5
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7680
      TabIndex        =   4
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7680
      TabIndex        =   3
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7680
      TabIndex        =   2
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7680
      TabIndex        =   1
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4185
      TabIndex        =   0
      Top             =   450
      Width           =   360
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private bPuedeMagia As Boolean
Private bPuedeMeditar As Boolean
Private bPuedeEscudo As Boolean
Private bPuedeCombateDistancia As Boolean

Private vsHelp(1 To NUMSKILLS) As String

Private Sub Form_Load()
    
    MirandoAsignarSkills = True
    
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Flags para saber que skills se modificaron
    ReDim flags(1 To NUMSKILLS)
    
    Call ValidarSkills
    
    Me.Picture = General_Set_GUI("VentanaSkills")
    imgAceptar.Init s_Large
    imgCancelar.Init s_Large
    
    Dim i As Integer
        For i = 1 To 20
            imgMenos(i).Init s_Small, "-"
            imgMas(i).Init s_Small, "+"
        Next i
            
    Call LoadHelp
End Sub


Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
    If Alocados > 0 Then

        If Val(Text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) + 1
            flags(SkillIndex) = flags(SkillIndex) + 1
            Alocados = Alocados - 1
        End If
            
    End If
    
    puntos.Caption = Alocados
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
    If Alocados < SkillPoints Then
        
        If Val(Text1(SkillIndex).Caption) > 0 And flags(SkillIndex) > 0 Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) - 1
            flags(SkillIndex) = flags(SkillIndex) - 1
            Alocados = Alocados + 1
        End If
    End If
    
    puntos.Caption = Alocados
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoAsignarSkills = False
End Sub

Private Sub imgAceptar_Click()
    Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    Call WriteModifySkills(skillChanges())
    
    SkillPoints = Alocados
        frmMain.imgAsignarSkill.Caption = Alocados
    Unload Me
End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Apuñalar)
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Carpinteria)
End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Armas)
End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Proyectiles)
End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Wrestling)
End Sub

Private Sub imgComercio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Comerciar)
End Sub

Private Sub imgDomar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Domar)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Defensa)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Tacticas)
End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.herreria)
End Sub

Private Sub imgLiderazgo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Liderazgo)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Magia)
End Sub

Private Sub imgMas_Click(Index As Integer)
    Call SumarSkillPoint(Index)
End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Meditar)
End Sub



Private Sub LoadHelp()
    
    vsHelp(eSkill.Magia) = "Magia:" & vbCrLf & _
                            "- Representa la habilidad de un personaje de las áreas mágica." & vbCrLf & _
                            "- Indica la variedad de hechizos que es capaz de dominar el personaje."
    If Not bPuedeMagia Then
        vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If
    
    vsHelp(eSkill.Robar) = "Robar:" & vbCrLf & _
                            "- Habilidades de hurto. Nunca por medio de la violencia." & vbCrLf & _
                            "- Indica la probabilidad de éxito del personaje al intentar apoderarse de oro de otro, en caso de ser Ladrón, tambien podrá apoderarse de items."
    
    vsHelp(eSkill.Tacticas) = "Evasión en Combate:" & vbCrLf & _
                                "- Representa la habilidad general para moverse en combate entre golpes enemigos sin morir o tropezar en el intento." & vbCrLf & _
                                "- Indica la posibilidad de evadir un golpe físico del personaje."
    
    vsHelp(eSkill.Armas) = "Combate con Armas:" & vbCrLf & _
                            "- Representa la habilidad del personaje para manejar armas de combate cuerpo a cuerpo." & vbCrLf & _
                            "- Indica la probabilidad de impactar al oponente con armas cuerpo a cuerpo."
    
    vsHelp(eSkill.Meditar) = "Meditar:" & vbCrLf & _
                                "- Representa la capacidad del personaje de concentrarse para abstrarse dentro de su mente, y así revitalizar su fuerza espiritual." & vbCrLf & _
                                "- Indica la velocidad a la que el personaje recupera maná (Clases mágicas)."
    
    If Not bPuedeMeditar Then
        vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Apuñalar) = "Apuñalar:" & vbCrLf & _
                                "- Representa la destreza para inflingir daño grave con armas cortas." & vbCrLf & _
                                "- Indica la posibilidad de apuñalar al enemigo en un ataque. El Asesino es la única clase que no necesitará 10 skills para comenzar a entrenar esta habilidad."

    vsHelp(eSkill.Ocultarse) = "Ocultarse:" & vbCrLf & _
                                "- La habilidad propia de un personaje para mimetizarse con el medio y evitar se perciba su presencia." & vbCrLf & _
                                "- Indica la facilidad con la que uno puede desaparecer de la vista de los demás y por cuanto tiempo."
    
    vsHelp(eSkill.Supervivencia) = "Superivencia:" & vbCrLf & _
                                    "- Es el conjunto de habilidades necesarias para sobrevivir fuera de una ciudad en base a lo que la naturaleza ofrece." & vbCrLf & _
                                    "- Permite conocer la salud de las criaturas guiándose exclusivamente por su aspecto, así como encender fogatas junto a las que descansar."
    
    vsHelp(eSkill.Talar) = "Talar:" & vbCrLf & _
                            "- Es la habilidad en el uso del hacha para evitar desperdiciar leña y maximizar la efectividad de cada golpe dado." & vbCrLf & _
                            "- Indica la probabilidad de obtener leña por golpe."
    
    vsHelp(eSkill.Comerciar) = "Comercio:" & vbCrLf & _
                                "- Es la habilidad para regatear los precios exigidos en la compra y evitar ser regateado al vender." & vbCrLf & _
                                "- Indica que tan caro se compra en el comercio con NPCs."
    
    vsHelp(eSkill.Defensa) = "Defensa con Escudos:" & vbCrLf & _
                                "- Es la habilidad de interponer correctamente el escudo ante cada embate enemigo para evitar ser impactado sin perder el equilibrio y poder responder rápidamente con la otra mano." & vbCrLf & _
                                "- Indica las probabilidades de bloquear un impacto con el escudo."
    
    If Not bPuedeEscudo Then
        vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If


    vsHelp(eSkill.Pesca) = "Pesca:" & vbCrLf & _
                            "- Es el conjunto de conocimientos básicos para poder armar un señuelo, poner la carnada en el anzuelo y saber dónde buscar peces." & vbCrLf & _
                            "- Indica la probabilidad de tener éxito en cada intento de pescar."
    
    vsHelp(eSkill.Mineria) = "Minería:" & vbCrLf & _
                                "- Es el conjunto de conocimientos sobre los distintos minerales, el dónde se obtienen, cómo deben ser extraídos y trabajados." & vbCrLf & _
                                "- Indica la probabilidad de tener éxito en cada intento de minar y la capacidad, o no de convertir estos minerales en lingotes."
    
    vsHelp(eSkill.Carpinteria) = "Carpintería:" & vbCrLf & _
                                    "- Es el conjunto de conocimientos para saber serruchar, lijar, encolar y clavar madera con un buen nivel de terminación." & vbCrLf & _
                                    "- Indica la habilidad en el manejo de estas herramientas, el que tan bueno se es en el oficio de carpintero."
    
    vsHelp(eSkill.herreria) = "Herrería:" & vbCrLf & _
                                "- Es el conjunto de conocimientos para saber procesar cada tipo de mineral para fundirlo, forjarlo y crear aleaciones." & vbCrLf & _
                                "- Indica la habilidad en el manejo de estas técnicas, el que tan bueno se es en el oficio de herrero."
    
    vsHelp(eSkill.Liderazgo) = "Liderazgo:" & vbCrLf & _
                                "- Es la habilidad propia del personaje para convencer a otros a seguirlo en batalla." & vbCrLf & _
                                "- Permite crear clanes y partys"
    
    vsHelp(eSkill.Domar) = "Domar Animales:" & vbCrLf & _
                                "- Es la habilidad en el trato con animales para que estos te sigan y ayuden en combate." & vbCrLf & _
                                "- Indica la posibilidad de lograr domar a una criatura y qué clases de criaturas se puede domar."
    
    vsHelp(eSkill.Proyectiles) = "Combate a distancia:" & vbCrLf & _
                                "- Es el manejo de las armas de largo alcance." & vbCrLf & _
                                "- Indica la probabilidad de éxito para impactar a un enemigo con este tipo de armas."
    
    If Not bPuedeCombateDistancia Then
        vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & _
                                "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Wrestling) = "Combate sin armas:" & vbCrLf & _
                                "- Es la habilidad del personaje para entrar en combate sin arma alguna salvo sus propios brazos." & vbCrLf & _
                                "- Indica la probabilidad de éxito para impactar a un enemigo estando desarmado. El Bandido y Ladrón tienen habilidades extras asociadas a esta habilidad."
    
    vsHelp(eSkill.Navegacion) = "Navegación:" & vbCrLf & _
                                "- Es la habilidad para controlar barcos en el mar sin naufragar." & vbCrLf & _
                                "- Indica que clase de barcos se pueden utilizar."
    
End Sub

Private Sub imgMenos_Click(Index As Integer)
    Call RestarSkillPoint(Index)
End Sub

Private Sub imgMineria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Mineria)
End Sub

Private Sub imgNavegacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Navegacion)
End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Ocultarse)
End Sub

Private Sub imgPesca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Pesca)
End Sub

Private Sub imgRobar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Robar)
End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Supervivencia)
End Sub

Private Sub imgTalar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Talar)
End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
    lblHelp.Caption = vsHelp(eeSkill)
End Sub

Private Sub ValidarSkills()

    bPuedeMagia = True
    bPuedeMeditar = True
    bPuedeEscudo = True
    bPuedeCombateDistancia = True

    Select Case UserClase
        Case eClass.Warrior, eClass.Hunter, eClass.Worker, eClass.Thief
            bPuedeMagia = False
            bPuedeMeditar = False
        
        Case eClass.Pirat
            bPuedeMagia = False
            bPuedeMeditar = False
            bPuedeEscudo = False
        
        Case eClass.Mage, eClass.Druid
            bPuedeEscudo = False
            bPuedeCombateDistancia = False
            
    End Select
    
    '   Magia
    imgMas(1).Visible = bPuedeMagia
    imgMenos(1).Visible = bPuedeMagia

    '   Meditar
    imgMas(5).Visible = bPuedeMeditar
    imgMenos(5).Visible = bPuedeMeditar

    '   Escudos
    imgMas(11).Visible = bPuedeEscudo
    imgMenos(11).Visible = bPuedeEscudo

    '   Proyectiles
    imgMas(18).Visible = bPuedeCombateDistancia
    imgMenos(18).Visible = bPuedeCombateDistancia
End Sub

