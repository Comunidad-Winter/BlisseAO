VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argoth - Configuración del juego"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "General:"
      Height          =   7695
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   9255
      Begin VB.Frame Frame1 
         Caption         =   "Otras opciones"
         Height          =   2895
         Index           =   1
         Left            =   0
         TabIndex        =   92
         Top             =   4440
         Width           =   9135
         Begin VB.CheckBox Check9 
            Caption         =   "Utilizar Drag&&Drop de los items del inventario con el Click Derecho"
            Height          =   315
            Left            =   240
            TabIndex        =   99
            Top             =   960
            Width           =   6255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Configurar el mapeado de teclas"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   98
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Configuración del Mapa"
         Height          =   1290
         Left            =   0
         TabIndex        =   16
         Top             =   3000
         Width           =   9135
         Begin VB.CommandButton Command3 
            Caption         =   "Mostrar Mapa del Mundo"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   94
            Top             =   720
            Width           =   3975
         End
         Begin VB.CheckBox ChkMap 
            Caption         =   "Generar y mostrar minimapa"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   360
            Width           =   4725
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dialogos y Consola"
         Height          =   2925
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   9135
         Begin VB.CheckBox Check13 
            Caption         =   "Motrar Noticias de clan al conectarse"
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   1320
            Width           =   6015
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Ejecutar Argoth Console History"
            Height          =   375
            Left            =   360
            TabIndex        =   97
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CheckBox Check35 
            Caption         =   "Almacenar la consola en un archivo de historial (Puede acceder a leer y borrar el mismo desde el Argoth Console History)"
            Height          =   615
            Left            =   120
            TabIndex        =   96
            Top             =   1800
            Width           =   8055
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Almacenar los dialogos del área en la consola Principal"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1560
            Width           =   5055
         End
         Begin VB.TextBox txtCantMensajes 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   300
            Left            =   6000
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "5"
            Top             =   960
            Width           =   570
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Mostrar los dialogos del clan en la Pantalla Principal"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   6255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Mostrar los dialogos del clan en la Consola Principal"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   7455
         End
         Begin VB.Label Label13 
            Caption         =   "Cantidad de dialogos máximos a mostrar en la pantalla principal: "
            Height          =   255
            Left            =   360
            TabIndex        =   93
            Top             =   960
            Width           =   6375
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   7695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame14 
         Caption         =   "Sonidos FXs"
         Height          =   1935
         Left            =   0
         TabIndex        =   87
         Top             =   2280
         Width           =   9135
         Begin VB.CheckBox Check15 
            Caption         =   "Escuchar sonidos y efectos FXs"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Width           =   3135
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Habilitar sonidos con efecto 3D"
            Height          =   195
            Left            =   4080
            TabIndex        =   88
            Top             =   360
            Width           =   3975
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   675
            Index           =   1
            Left            =   120
            TabIndex        =   90
            Top             =   1080
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1191
            _Version        =   393216
            Max             =   100
            TickStyle       =   2
            TextPosition    =   1
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Volumen"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   91
            Top             =   720
            Width           =   8415
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Game4Fun PlayMusic Plugin"
         Height          =   3375
         Left            =   0
         TabIndex        =   85
         Top             =   4320
         Width           =   9135
         Begin VB.Label Label12 
            Caption         =   "Debe tener instalado el Game4Fun PlayMusic para utilizar esta función."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   735
            Left            =   240
            TabIndex        =   86
            Top             =   360
            Width           =   7815
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Música de ambiente"
         Height          =   2175
         Left            =   0
         TabIndex        =   80
         Top             =   0
         Width           =   9135
         Begin VB.CheckBox Check16 
            Caption         =   "Escuchar música de ambiente"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Detener Música"
            Height          =   375
            Left            =   240
            TabIndex        =   81
            Top             =   720
            Width           =   1695
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   1320
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1191
            _Version        =   393216
            Max             =   100
            TickStyle       =   2
            TextPosition    =   1
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Volumen"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   84
            Top             =   960
            Width           =   8535
         End
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   7695
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame10 
         Caption         =   "Links útiles"
         Height          =   1215
         Left            =   0
         TabIndex        =   72
         Top             =   3600
         Width           =   9135
         Begin VB.CommandButton Command3 
            Caption         =   "Foro Oficial"
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   75
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Web Oficial"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ayuda general"
         Height          =   3375
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   9135
         Begin VB.CommandButton Command3 
            Caption         =   "Wiki Argoth"
            Height          =   375
            Index           =   6
            Left            =   5520
            TabIndex        =   73
            Top             =   1200
            Width           =   3495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Manual Oficial"
            Height          =   375
            Index           =   3
            Left            =   5520
            TabIndex        =   71
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Soporte"
            Height          =   375
            Index           =   5
            Left            =   5520
            TabIndex        =   70
            Top             =   2640
            Width           =   3495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Mostrar Tutorial"
            Height          =   375
            Index           =   0
            Left            =   5520
            TabIndex        =   69
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Label Label11 
            Caption         =   "Para conocer un poco más sobre la jugabilidad y la interface de Argoth, podrás ejecutar este Tutorial Interactivo."
            Height          =   495
            Left            =   120
            TabIndex        =   79
            Top             =   1920
            Width           =   5415
         End
         Begin VB.Label Label10 
            Caption         =   $"frmOpciones.frx":000C
            Height          =   615
            Left            =   120
            TabIndex        =   78
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Label Label9 
            Caption         =   $"frmOpciones.frx":00A2
            Height          =   855
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   5295
         End
         Begin VB.Label Label8 
            Caption         =   "El manual oficial del juego te orientará rápidamente sobre la jugabilidad principal de Argoth."
            Height          =   495
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   5415
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame15 
         Caption         =   "Configuración automática:"
         Height          =   3135
         Left            =   5040
         TabIndex        =   56
         Top             =   4560
         Width           =   4215
         Begin VB.TextBox txtInfoSet 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            ToolTipText     =   "Información"
            Top             =   1320
            Width           =   3975
         End
         Begin MSComctlLib.Slider Slider2 
            Height          =   495
            Left            =   360
            TabIndex        =   57
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   873
            _Version        =   393216
            Min             =   1
            Max             =   4
            SelStart        =   3
            Value           =   3
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   62
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Máxima"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   61
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Media"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   60
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Mínima"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   59
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Graficos"
         Height          =   4455
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   9135
         Begin VB.CheckBox Check34 
            Caption         =   "Calcular TileBuffer de manera automática (Recomendado)"
            Height          =   435
            Left            =   3960
            TabIndex        =   65
            Top             =   2880
            Width           =   4815
         End
         Begin VB.CheckBox Check21 
            Caption         =   "Mostrar daño de ataque"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2160
            Width           =   2895
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Utilizar Sombras de baja calidad"
            Height          =   255
            Left            =   3960
            TabIndex        =   63
            Top             =   720
            Width           =   3255
         End
         Begin MSComctlLib.Slider HScroll1 
            Height          =   435
            Left            =   5040
            TabIndex        =   49
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   767
            _Version        =   393216
            Min             =   6
            Max             =   12
            SelStart        =   8
            Value           =   8
         End
         Begin VB.CheckBox Check32 
            Caption         =   "Utilizar desvanecimiento en los textos"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   3240
            Width           =   3735
         End
         Begin VB.CheckBox Check31 
            Caption         =   "Utilizar desvanecimiento en los techos"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   3000
            Width           =   3735
         End
         Begin VB.CheckBox Check30 
            Caption         =   "Permitir desenfoque en movimiento"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3960
            TabIndex        =   45
            Top             =   1560
            Width           =   3375
         End
         Begin VB.CheckBox Check29 
            Caption         =   "Mostrar daño en el mapa"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CheckBox Check28 
            Caption         =   "Mostrar huellas de los personajes"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   2640
            Width           =   3255
         End
         Begin VB.CheckBox Check27 
            Caption         =   "Pantalla de ingreso animada"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3960
            TabIndex        =   42
            Top             =   3960
            Width           =   4095
         End
         Begin VB.CheckBox Check26 
            Caption         =   "Mostrar sangre en el suelo"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CheckBox Check25 
            Caption         =   "Mostrar MiniMapa Zonal"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3600
            Width           =   2415
         End
         Begin VB.CheckBox Check24 
            Caption         =   "Utilizar Sombras direccionales"
            Height          =   255
            Left            =   3960
            TabIndex        =   39
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox Check23 
            Caption         =   "Partículas en Mouse"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Mostrar el movimiento del agua"
            Height          =   255
            Left            =   3960
            TabIndex        =   37
            Top             =   1080
            Width           =   3015
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Calcular y Renderizar FPS"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1560
            Width           =   3495
         End
         Begin VB.CheckBox Check22 
            Caption         =   "Utilizar Partículas en la meditación"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox Check20 
            Caption         =   "Mostrar Sombras en Personajes"
            Height          =   255
            Left            =   3960
            TabIndex        =   28
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Señalar y tonalizar objetivos"
            Height          =   255
            Left            =   3960
            TabIndex        =   27
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Mostrar trayectoria de los proyectiles"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   3495
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Utilizar Partículas"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Mostrar información de los grupos en la pantalla"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   2520
            Width           =   4695
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Mostrar Nombre de los Items al pasar el Mouse"
            Height          =   255
            Left            =   3960
            TabIndex        =   23
            Top             =   1920
            Width           =   5055
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Renderizado dinámico del inventario"
            Height          =   255
            Left            =   3960
            TabIndex        =   22
            Top             =   1320
            Width           =   3495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Limitar FPS"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox Check33 
            Caption         =   "Utilizar efectos de Envenenamiento, Paralisis y Inmovilidad"
            Enabled         =   0   'False
            Height          =   555
            Left            =   120
            TabIndex        =   48
            Top             =   3840
            Width           =   3735
         End
         Begin VB.Label Label4 
            Caption         =   "Tile Buffer:"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   66
            Top             =   3360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Engine:"
         Height          =   3135
         Left            =   0
         TabIndex        =   2
         Top             =   4560
         Width           =   4935
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Caption         =   "Filtrado de gráficos:"
            Height          =   735
            Left            =   120
            TabIndex        =   50
            Top             =   2280
            Width           =   4455
            Begin VB.OptionButton Option4 
               Caption         =   "Antisópico (Automático)"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   55
               Top             =   480
               Width           =   3255
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Trilinear"
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   54
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Bilinear"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   53
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Ninguno"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Filtrar Texturas:"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   4455
            Begin VB.OptionButton Option6 
               Caption         =   "512MB"
               Height          =   255
               Index           =   3
               Left            =   3000
               TabIndex        =   36
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option6 
               Caption         =   "256MB"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   35
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option6 
               Caption         =   "128MB"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   34
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option6 
               Caption         =   "64MB"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Precarga máxima de texturas:"
               Height          =   255
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   3735
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Utilizar nuevo sistema de luces radiales"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Utilizar sincronización vertical"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Mixed"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Hardware"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Software"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ejecutar el juego en modo ventana"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Aceleración:"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "El cambio de estas opciones tendrá efecto al reiniciar el Cliente."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8175
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   14420
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Engine Gráfico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sonido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Información"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   6840
      TabIndex        =   58
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Menu mnusaveex 
      Caption         =   "Guardar y Salir"
   End
   Begin VB.Menu mnudefaults 
      Caption         =   "Restaurar Defaults"
   End
   Begin VB.Menu mnusave 
      Caption         =   "Salir sin Guardar"
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager


Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub Check1_Click()
    If Check1.value = Checked Then
        Settings.vSync = True
        Check3.Enabled = False
    Else
        Settings.vSync = False
        Check3.Enabled = True
    End If
End Sub









Private Sub Check10_Click()
    If Check10.value = Checked Then
        Settings.MostrarFPS = True
    Else
        Settings.MostrarFPS = False
    End If
End Sub

Private Sub Check12_Click()
    If Check12.value = Checked Then
        Settings.Water_Effect = True
    Else
        Settings.Water_Effect = False
    End If
End Sub

Private Sub Check13_Click()
    If Check13.value = Checked Then
        Settings.GuildNews = True
    Else
        Settings.GuildNews = False
    End If
End Sub

Private Sub Check14_Click()
    If loading Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)

    If Check14.value = Checked Then
        bSoundEffectsActivated = True
        Audio.SoundEffectsActivated = bSoundEffectsActivated
        Settings.Sonido3D = True
    Else
        bSoundEffectsActivated = False
        Audio.SoundEffectsActivated = bSoundEffectsActivated
        Settings.Sonido3D = False
    End If
End Sub

Private Sub Check15_Click()
    If loading Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    
    If Check15.value = Checked Then
        bSoundActivated = True
        Settings.Sonido = True
    Else
        bSoundActivated = False
        Settings.Sonido = False
    End If
    
    If Not bSoundActivated Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        Slider1(1).Enabled = False
    Else
        Audio.SoundActivated = True
        Slider1(1).Enabled = True
        Slider1(1).value = Settings.SoundVolume
    End If
End Sub

Private Sub Check16_Click()
    If loading Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    
    If Check16.value = Checked Then
        bMusicActivated = True
        Settings.Musica = True
    Else
        bMusicActivated = False
        Settings.Musica = False
        Audio.StopMidi
        Audio.MP3_Stop
    End If
            
    If Not bMusicActivated Then
        Audio.MusicActivated = False
        Slider1(0).Enabled = False
    Else
        If Not Audio.MusicActivated Then
            Audio.MusicActivated = True
            Slider1(0).Enabled = True
            Slider1(0).value = Audio.MusicVolume
        End If
    End If
End Sub

Private Sub Check17_Click()
    If Check17.value = Checked Then
        Settings.ParticleEngine = True
        Check22.Enabled = True
    Else
        Settings.ParticleEngine = False
        Settings.ParticleMeditation = False
        Check22.value = Unchecked
        Check22.Enabled = False
    End If
End Sub

Private Sub Check18_Click()
    If Check18.value = Checked Then
        Settings.ProyectileEngine = True
    Else
        Settings.ProyectileEngine = False
    End If
End Sub

Private Sub Check19_Click()
    If Check19.value = Checked Then
        Settings.TonalidadPJ = True
    Else
        Settings.TonalidadPJ = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.value = Checked Then
        Settings.Ventana = True
    Else
        Settings.Ventana = False
    End If
End Sub

Private Sub Check20_Click()
    If Check20.value = Checked Then
        Settings.UsarSombras = True
    Else
        Settings.UsarSombras = False
    End If
End Sub

Private Sub Check22_Click()
    If Check22.value = Checked Then
        Settings.ParticleMeditation = True
    Else
        Settings.ParticleMeditation = False
    End If
End Sub

Private Sub Check23_Click()
    If Check23.value = Checked Then
        Settings.Mouse_Effect = True
    Else
        Settings.Mouse_Effect = False
    End If
End Sub

Private Sub Check24_Click()
    If Check24.value = Checked Then
        Settings.Shadow_Effect = True
    Else
        Settings.Shadow_Effect = False
    End If
End Sub

Private Sub Check28_Click()
    If Check28.value = Checked Then
        Settings.Walk_Effect = True
    Else
        Settings.Walk_Effect = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.value = Checked Then
        Settings.LimiteFPS = True
    Else
        Settings.LimiteFPS = False
    End If
End Sub

Private Sub Check32_Click()
    If Check32.value = Checked Then
        Settings.Text_Effect = 1
    Else
        Settings.Text_Effect = 0
    End If
End Sub

Private Sub Check4_Click()
    If Check4.value = Checked Then
        Settings.Luces = 1
    Else
        Settings.Luces = 0
    End If
End Sub

Private Sub Check5_Click()
    If Check5.value = Checked Then
        Settings.DinamicInventory = True
    Else
        Settings.DinamicInventory = False
    End If
End Sub

Private Sub Check6_Click()
    If Check6.value = Checked Then
        Settings.NombreItems = True
    Else
        Settings.NombreItems = False
    End If
End Sub

Private Sub Check7_Click()
    If Check7.value = Checked Then
        Settings.PartyMembers = True
    Else
        Settings.PartyMembers = False
    End If
End Sub

Private Sub Check8_Click()
    If Check8.value = Checked Then
        Settings.DialogosEnConsola = True
    Else
        Settings.DialogosEnConsola = False
    End If
End Sub

Private Sub Check9_Click()
    If Check9.value = Checked Then
        Settings.DragerDrop = True
    Else
        Settings.DragerDrop = False
    End If
End Sub

Private Sub ChkMap_Click()
If ChkMap.value = Checked Then
    Settings.MiniMap = True
    Call General_Pixel_Map_Render
Else
    Settings.MiniMap = False
End If
End Sub


Private Sub Command2_Click(index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
        
    Select Case index
        Case 0
            Call frmCustomKeys.Show(vbModal, Me)
    End Select
End Sub

Private Sub Command3_Click(index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)

    Select Case index
        Case 1
            'Call frmMapa.Show(vbModal, Me)
        Case 2
            Call ShellExecute(0, "Open", Client_Forum, "", App.Path, SW_SHOWNORMAL)
        Case 3
            Call ShellExecute(0, "Open", Client_Web & "manual/", "", App.Path, SW_SHOWNORMAL)
        Case 4
            Call ShellExecute(0, "Open", Client_Web, "", App.Path, SW_SHOWNORMAL)
        Case 5
            Call ShellExecute(0, "Open", Client_Web & "soporte.php", "", App.Path, SW_SHOWNORMAL)
    End Select

End Sub

Private Sub Command4_Click()
    Audio.StopMidi
End Sub

Private Sub HScroll1_Change()
    Settings.BufferSize = HScroll1.value
    Engine_Set_TileBuffer Settings.BufferSize
End Sub

Private Sub mnudefaults_Click()
    If MsgBox("¿Estás seguro que deseas resetear las opciones a default?", vbYesNo) = vbYes Then
        Call Settings_Set_Default
        Call Settings_Save
    End If
End Sub

Private Sub mnuinfo_Click()
    Call ShellExecute(0, "Open", Client_Web & "manual.php?pagina=opciones_engine", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub mnusave_Click()
    Unload Me
End Sub

Private Sub mnusaveex_Click()
    Settings_Save
    Unload Me
    If frmMain.Visible And Not frmConnect.Visible Then frmMain.SetFocus
End Sub

Private Sub Option1_Click(index As Integer)
    '0-Software 1-Hardware 2-Mixed
    Settings.Aceleracion = index
    
    If index = 2 Or index = 1 Then
        MsgBox "ATENCIÓN: Los Modos Mixed o Hardware requieren tener instalada una placa de video, si usted posee video Onboard no active estas opciones.", vbCritical, "Blisse-AO: Engine Gráfico"
    End If
End Sub

Private Sub Option2_Click(index As Integer)
    Select Case index
        Case 0
            DialogosClanes.Activo = False
            Settings.DialogoClanesActivo = False
        Case 1
            DialogosClanes.Activo = True
            Settings.DialogoClanesActivo = True
    End Select
End Sub

Private Sub Option3_Click(index As Integer)
    Select Case index
        Case 0
            Settings.Dialog_Align = 0
        Case 1
            Settings.Dialog_Align = 1
    End Select
End Sub





Private Sub Option4_Click(index As Integer)
    Select Case index
        Case 0 '
            Engine_Set_Texture_Filter DirectDevice, DirectCaps, 0, TexFilter_None
        Case 1 '
            Engine_Set_Texture_Filter DirectDevice, DirectCaps, 0, TexFilter_Bilinear
        Case 2 '
            Engine_Set_Texture_Filter DirectDevice, DirectCaps, 0, TexFilter_Trilinear
        Case 3 '
            Engine_Set_Texture_Filter DirectDevice, DirectCaps, 0, TexFilter_Anisotropic, 16
    End Select
End Sub

Private Sub Option6_Click(index As Integer)
    Select Case index
        Case 0 '64MB
            Settings.MemoryVideoMax = 64
        Case 1 '128MB
            Settings.MemoryVideoMax = 128
        Case 2 '256MB
            Settings.MemoryVideoMax = 256
        Case 3 '512MB
            Settings.MemoryVideoMax = 512
    End Select
End Sub

Private Sub TabStrip1_Click()
    Frame2.Visible = False 'Graf
    Frame1(0).Visible = False 'Sound
    Frame11.Visible = False 'Info
    Frame4.Visible = False 'General
    
    Select Case TabStrip1.SelectedItem.index
        Case 1 'General
            Frame4.Visible = True 'General
        Case 2 'Engine Gráfico
            Frame2.Visible = True 'Graf
        Case 3
            Frame1(0).Visible = True 'Sound
        Case 4
            Frame11.Visible = True 'Info
    End Select
End Sub

Private Sub txtCantMensajes_Change()
    txtCantMensajes.Text = Val(txtCantMensajes.Text)
    
    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
        Settings.DialogoClanesCant = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5
        Settings.DialogoClanesCant = 5
    End If
End Sub


Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub LoadUserConfig()

    '   Load music config
    bMusicActivated = Audio.MusicActivated
    Slider1(0).Enabled = bMusicActivated
    
    If bMusicActivated Then
        Check16.value = Checked
        Slider1(0).value = Audio.MusicVolume
        
    End If
    
    '   Load Sound config
    bSoundActivated = Audio.SoundActivated
    Slider1(1).Enabled = bSoundActivated
    
    If bSoundActivated Then
        Check15.value = Checked
        Slider1(1).value = Settings.SoundVolume
    End If
    
    '   Load Sound Effects config
    bSoundEffectsActivated = Audio.SoundEffectsActivated
    If bSoundEffectsActivated Then Check14.value = Checked
    
    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
    If DialogosClanes.Activo Then
        Option2(1).value = True
    Else
        Option2(0).value = True
    End If
    
    If Settings.Dialog_Align = 1 Then
    '    Option3(1).value = True
    Else
    '    Option3(0).value = True
    End If
    
    If Settings.GuildNews Then
        Check13.value = Checked
    End If
        
    '   Engine Settings
    Option1(Settings.Aceleracion).value = True
    If Settings.vSync Then
        Check1.value = Checked
        Check3.Enabled = False
    End If
    If Settings.Ventana Then Check2.value = Checked
    If Settings.Luces = 1 Then Check4.value = Checked '1 redondas, 0 cuadradas ¬¬
    If Settings.DinamicInventory Then Check5.value = Checked
    If Settings.NombreItems Then Check6.value = Checked
    If Settings.PartyMembers Then Check7.value = Checked
    If Settings.LimiteFPS Then Check3.value = Checked
    If Settings.ParticleEngine Then Check17.value = Checked
    If Settings.ProyectileEngine Then Check18.value = Checked
    If Settings.TonalidadPJ Then Check19.value = Checked
    If Settings.UsarSombras Then Check20.value = Checked
    If Settings.ParticleMeditation Then Check22.value = Checked
    If Settings.MostrarFPS Then Check10.value = Checked
    
    If Settings.Water_Effect Then Check12.value = Checked
    If Settings.Mouse_Effect Then Check23.value = Checked
    
    If Settings.Shadow_Effect Then Check24.value = Checked
    If Settings.Walk_Effect Then Check28.value = Checked
    If Settings.Text_Effect Then Check32.value = Checked
    
    HScroll1.value = Settings.BufferSize
    
    '   Settings
    If Settings.DialogosEnConsola Then Check8.value = Checked
    If Settings.DragerDrop Then Check9.value = Checked
   
    Select Case Settings.MemoryVideoMax
        Case 64
            Option6(0).value = True
        Case 128
            Option6(1).value = True
        Case 256
            Option6(2).value = True
        Case 512
            Option6(3).value = True
    End Select
    
    'MiniMap [TonchitoZ]
    If Settings.MiniMap Then
        ChkMap.value = Checked
    Else
        ChkMap.value = Unchecked
    End If
    
End Sub

Private Sub Slider1_Change(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
            Settings.SoundVolume = Slider1(1).value
    End Select
End Sub

Private Sub Slider1_Scroll(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
            Settings.SoundVolume = Slider1(1).value
    End Select
End Sub
