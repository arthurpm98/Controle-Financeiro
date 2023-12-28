VERSION 5.00
Object = "{82392BA0-C18D-11D2-B0EA-00A024695830}#1.0#0"; "ticaldr6.ocx"
Begin VB.Form FormReceitas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Receitas"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3540
      TabIndex        =   19
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdExcluir 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   4
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10725
      TabIndex        =   3
      Top             =   8505
      Width           =   1125
   End
   Begin VB.CommandButton cmdEditar 
      Appearance      =   0  'Flat
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   135
      TabIndex        =   2
      Top             =   8505
      Width           =   1110
   End
   Begin VB.CommandButton cmdGravar 
      Appearance      =   0  'Flat
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   8505
      Width           =   1125
   End
   Begin VB.Frame FraReceitas 
      Appearance      =   0  'Flat
      Caption         =   "Receitas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   8040
      Left            =   150
      TabIndex        =   0
      Top             =   405
      Width           =   11700
      Begin VB.OptionButton OptPendente 
         Appearance      =   0  'Flat
         Caption         =   "Pendente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6240
         TabIndex        =   18
         Top             =   3765
         Width           =   1020
      End
      Begin VB.OptionButton OptPago 
         Appearance      =   0  'Flat
         Caption         =   "Pago"
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5400
         TabIndex        =   17
         Top             =   3645
         Width           =   735
      End
      Begin VB.TextBox txtValor 
         Height          =   345
         Left            =   5430
         TabIndex        =   16
         Top             =   2595
         Width           =   1815
      End
      Begin VB.TextBox txtObservacao 
         Height          =   285
         Left            =   1275
         TabIndex        =   14
         Top             =   1500
         Width           =   9495
      End
      Begin VB.ComboBox ComCategoria 
         Height          =   315
         Left            =   5430
         TabIndex        =   12
         Top             =   2100
         Width           =   5355
      End
      Begin TDBCalendar6Ctl.TDBCalendar TDBDataReceita 
         Height          =   2175
         Left            =   300
         TabIndex        =   10
         Top             =   2475
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   3836
         ShowContextMenu =   -1  'True
         Appearance      =   1
         AutoSize        =   0   'False
         BorderStyle     =   1
         BackColor       =   -2147483643
         StartOfMonth    =   0
         EmptyRows       =   0
         Enabled         =   -1  'True
         FirstMonth      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LineColors0     =   -2147483632
         LineStyles0     =   0
         LineColors1     =   -2147483632
         LineStyles1     =   0
         LineColors2     =   -2147483632
         LineStyles2     =   0
         LineColors3     =   -2147483632
         LineStyles3     =   0
         LineColors4     =   -2147483632
         LineStyles4     =   0
         LineColors5     =   -2147483632
         LineStyles5     =   0
         LineColors6     =   -2147483632
         LineStyles6     =   2
         MarginBottom    =   0
         MarginTitle     =   0
         MarginTop       =   0
         MarginLeft      =   0
         MarginRight     =   0
         MarginWidth     =   0
         MarginHeight    =   0
         MaxDate         =   5373484
         MinDate         =   1757585
         MousePointer    =   0
         YearType        =   0
         MonthRows       =   1
         MonthCols       =   1
         MultiSelect     =   0
         NavOrientation  =   2
         ScrollRate      =   1
         ScrollTipAlign  =   0
         SelEdgeWidth    =   8
         SelectStyle     =   0
         SelectWhat      =   0
         ShowMenu        =   -1  'True
         ShowNavigator   =   3
         ShowScrollTip   =   -1  'True
         ShowTrailing    =   -1  'True
         StartOfWeek     =   1
         Templates       =   0
         TipInterval     =   500
         TitleHeight     =   0
         TitleFormat     =   "mmmm yyy"
         ValueIsNull     =   0   'False
         Value           =   2460096
         OverrideTipText =   ""
         TopDate         =   2460066
         AttribStyles    =   "FormReceitas.frx":0000
         StyleSets       =   "FormReceitas.frx":00C0
         CtrlType        =   8
         CtrlValue       =   "CtrlStyle"
         DayType         =   8
         DayValue        =   "DayStyle"
         TitleType       =   8
         TitleValue      =   "TitleStyle"
         WeekType        =   8
         WeekValue       =   "WeekStyle"
         TrailType       =   8
         TrailValue      =   "TrailAttrib"
         SelType         =   8
         SelValue        =   "SelAttrib"
         WeekRests0      =   0
         WeekReflect0    =   0
         WeekCaption0    =   "dom"
         WeekAttrib0Type =   8
         WeekAttrib0Value=   "SunAttrib"
         WeekRests1      =   0
         WeekReflect1    =   0
         WeekCaption1    =   "seg"
         WeekAttrib1Type =   1
         WeekRests2      =   0
         WeekReflect2    =   0
         WeekCaption2    =   "ter"
         WeekAttrib2Type =   1
         WeekRests3      =   0
         WeekReflect3    =   0
         WeekCaption3    =   "qua"
         WeekAttrib3Type =   1
         WeekRests4      =   0
         WeekReflect4    =   0
         WeekCaption4    =   "qui"
         WeekAttrib4Type =   1
         WeekRests5      =   0
         WeekReflect5    =   0
         WeekCaption5    =   "sex"
         WeekAttrib5Type =   1
         WeekRests6      =   0
         WeekReflect6    =   0
         WeekCaption6    =   "sáb"
         WeekAttrib6Type =   8
         WeekAttrib6Value=   "SatAttrib"
         HolidayStyles   =   "FormReceitas.frx":0228
         UserStyles      =   ""
         Key             =   "FormReceitas.frx":0244
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1290
         TabIndex        =   8
         Top             =   990
         Width           =   9480
      End
      Begin VB.TextBox txtCodigo 
         Height          =   300
         Left            =   1305
         TabIndex        =   6
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4935
         TabIndex        =   15
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label lblObservacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   285
         TabIndex        =   13
         Top             =   1515
         Width           =   945
      End
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4635
         TabIndex        =   11
         Top             =   2130
         Width           =   735
      End
      Begin VB.Label lblDataReceita 
         AutoSize        =   -1  'True
         Caption         =   "Data do Pagamento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   9
         Top             =   2130
         Width           =   1440
      End
      Begin VB.Label lblDescrição 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   7
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label lblCódigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   690
         TabIndex        =   5
         Top             =   510
         Width           =   540
      End
   End
   Begin VB.Menu mnuReceitas 
      Caption         =   "Receitas"
      Index           =   0
      Begin VB.Menu mnuReceitasNova 
         Caption         =   "Nova Receita"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FormReceitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

