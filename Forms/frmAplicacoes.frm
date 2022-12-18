VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmAplicacoes 
   Appearance      =   0  'Flat
   Caption         =   "Aplicações"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   20385
   Icon            =   "frmAplicacoes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   20385
   WindowState     =   2  'Maximized
   Begin VB.Frame quadBotoes 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   10800
      Left            =   19440
      TabIndex        =   13
      Top             =   -120
      Width           =   1335
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":680A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":7AD6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":8DA5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   2160
         Width           =   915
         _Version        =   131072
         _ExtentX        =   1614
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":A071
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":B33F
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":C60B
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":D8D6
      End
   End
   Begin VB.Frame quadRodape 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   10200
      Width           =   19455
      Begin VB.Label lblRodape 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         TabIndex        =   29
         Top             =   80
         Width           =   11895
      End
   End
   Begin VB.Frame quadCadastro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   10215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19455
      Begin VB.ComboBox cmbAno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   44
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbCorretora 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   34
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkInvestir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Investir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   33
         Top             =   885
         Width           =   1215
      End
      Begin VB.Frame quadInvestimento 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Investimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   360
         TabIndex        =   31
         Top             =   1800
         Width           =   18615
         Begin EditLib.fpCurrency txtSaque 
            Height          =   375
            Left            =   6525
            TabIndex        =   47
            Top             =   5880
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
            _ExtentY        =   661
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "R$ 0,00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpCurrency txtValor 
            Height          =   375
            Left            =   525
            TabIndex        =   46
            Top             =   5880
            Width           =   1935
            _Version        =   196608
            _ExtentX        =   3413
            _ExtentY        =   661
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "R$ 0,00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin FPSpreadADO.fpSpread gridSimulacao 
            Height          =   4845
            Left            =   8280
            TabIndex        =   45
            Top             =   360
            Width           =   4395
            _Version        =   458752
            _ExtentX        =   7752
            _ExtentY        =   8546
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   14737632
            GridColor       =   8421504
            MaxCols         =   3
            MaxRows         =   100
            OperationMode   =   1
            ShadowColor     =   12632256
            ShadowDark      =   8421504
            ShadowText      =   0
            SpreadDesigner  =   "frmAplicacoes.frx":EBA1
            UserResize      =   0
         End
         Begin VB.TextBox txtTaxa 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4995
            TabIndex        =   8
            Top             =   5880
            Width           =   735
         End
         Begin VB.ComboBox cmbMes 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2835
            TabIndex        =   7
            Top             =   5880
            Width           =   1815
         End
         Begin FPSpreadADO.fpSpread gridAplicacoes 
            Height          =   4965
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   7755
            _Version        =   458752
            _ExtentX        =   13679
            _ExtentY        =   8758
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   14737632
            GridColor       =   8421504
            MaxCols         =   5
            MaxRows         =   100
            OperationMode   =   1
            ShadowColor     =   12632256
            ShadowDark      =   8421504
            ShadowText      =   0
            SpreadDesigner  =   "frmAplicacoes.frx":F95E
            UserResize      =   0
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdLimparAplicacoes 
            Height          =   1095
            Left            =   17235
            TabIndex        =   10
            Top             =   5400
            Width           =   975
            _Version        =   131072
            _ExtentX        =   1720
            _ExtentY        =   1931
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   0   'False
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   0   'False
            BackStyle       =   1
            AutoSize        =   0
            AutoSizeOffsetTop=   0
            AutoSizeOffsetBottom=   0
            AutoSizeOffsetLeft=   0
            AutoSizeOffsetRight=   0
            DropShadowOffsetX=   3
            DropShadowOffsetY=   3
            DropShadowType  =   0
            DropShadowColor =   0
            Redraw          =   -1  'True
            ButtonDesigner  =   "frmAplicacoes.frx":12F4F
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdAplicar 
            Height          =   945
            Left            =   16155
            TabIndex        =   9
            Top             =   5400
            Width           =   975
            _Version        =   131072
            _ExtentX        =   1720
            _ExtentY        =   1667
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   0   'False
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   0   'False
            BackStyle       =   1
            AutoSize        =   0
            AutoSizeOffsetTop=   0
            AutoSizeOffsetBottom=   0
            AutoSizeOffsetLeft=   0
            AutoSizeOffsetRight=   0
            DropShadowOffsetX=   3
            DropShadowOffsetY=   3
            DropShadowType  =   0
            DropShadowColor =   0
            Redraw          =   -1  'True
            ButtonDesigner  =   "frmAplicacoes.frx":14228
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   555
            TabIndex        =   43
            Top             =   5640
            Width           =   495
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mês"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2835
            TabIndex        =   42
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Taxa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   4995
            TabIndex        =   41
            Top             =   5640
            Width           =   735
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   6555
            TabIndex        =   40
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "R$"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   39
            Top             =   5925
            Width           =   255
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "R$"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   6255
            TabIndex        =   38
            Top             =   5925
            Width           =   255
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   5790
            TabIndex        =   37
            Top             =   5925
            Width           =   255
         End
      End
      Begin VB.Frame quadDatas 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   975
         Left            =   15720
         TabIndex        =   21
         Top             =   8880
         Width           =   3255
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAtualizacao 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cadastro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Atualização"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkInativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I&nativo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12135
         TabIndex        =   6
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Corretora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   4680
         TabIndex        =   35
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   7200
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame quadPesquisa 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19455
      Begin VB.CheckBox chkInativoPesquisa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I&nativos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   17400
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPesquisa 
         Height          =   495
         Left            =   7800
         TabIndex        =   26
         Top             =   360
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   873
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   0
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmAplicacoes.frx":154F8
      End
      Begin VB.TextBox txtPesquisa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   7335
      End
      Begin FPSpreadADO.fpSpread gridPrincipal 
         Height          =   8925
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   18705
         _Version        =   458752
         _ExtentX        =   32994
         _ExtentY        =   15743
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   7
         MaxRows         =   1
         OperationMode   =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmAplicacoes.frx":167CF
         UserResize      =   0
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   260
         TabIndex        =   27
         Top             =   200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fOpcao As Integer
Dim fClsAplicacoes As New ClsAplicacoes
Dim fCodigo As Integer
Dim fCondicao As String
Dim fClsExtrato As New clsExtrato
Private Enum EnumGrid
    eCodigo = 1
    eNome
    eCadastro
    eAtualizacao
    eStatus
End Enum

Private Sub AlimentarGrid()
Dim sSql As String

sSql = "SELECT APL_CODIGO AS CODIGO, APL_NOME AS NOME, COR_NOME AS CORRETORA, APL_ANO AS ANO, APL_CADASTRO AS CADASTRO, APL_ATUALIZACAO AS ATUALIZAÇÃO, " _
    & " CASE WHEN APL_INATIVO = 0 THEN 'ATIVO'" _
    & " WHEN APL_INATIVO = 1 THEN 'INATIVO' END AS STATUS FROM APLICACOES INNER JOIN CORRETORAS ON COR_CODIGO = APL_CORRETORA" _
    & " WHERE 1 = 1 "

If chkInativoPesquisa Then
    sSql = sSql & " AND APL_INATIVO = 1"
Else
    sSql = sSql & " AND APL_INATIVO = 0"
End If

If txtPesquisa.Text <> "" Then
    sSql = sSql & " AND APL_NOME LIKE '" & Trim(txtPesquisa.Text) & "%'"
End If

SpreadClean gridPrincipal
SpreadFill gridPrincipal, sSql

Exit Sub
End Sub

Private Sub chkInvestir_Click()
    If chkInvestir.Value Then
        quadInvestimento.Visible = True
        chkInvestir.Enabled = False
        SpreadClean gridAplicacoes
        SpreadClean gridSimulacao
    End If
End Sub

Private Sub cmdAplicar_Click()
On Error GoTo vbErrorHandler
Dim sMes As Long, sLinhas As Long, sCont As Long, sValor As Double

If Not AnalisarDados Then Exit Sub

With gridAplicacoes
    For sCont = 1 To .MaxRows
        .Row = sCont
        .col = 2
        If .Text <> "" Then
            sMes = RetornaNumeroMes(.Text)
            If sMes > cmbMes.ListIndex Then
                MsgBox "Não é possível inserir uma aplicação retroativa!", vbInformation, "GoInvest"
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next
End With

If cmbMes.ListIndex <= CInt(Month(Date)) Then
    fClsExtrato.Saldo = 0
    With gridAplicacoes
        For sCont = 1 To .MaxRows
            .Row = sCont
            .col = 1
            If .Text = 0 Then
                .Row = sCont
                .RowHidden = False
                .SetText 1, sCont, cmbMes.ListIndex & Second(Time) & Day(Date)
                .SetText 2, sCont, MonthName(cmbMes.ListIndex)
                .SetText 3, sCont, CDbl(txtValor)
                .SetText 4, sCont, CDbl(txtTaxa)
                .SetText 5, sCont, CDbl(txtSaque)
                Exit For
            End If
        Next
    End With
    fClsExtrato.Taxa = CDbl(txtTaxa.Text)
    SomarValorMensal
    For sMes = (cmbMes.ListIndex) To 12
        With gridSimulacao
            sValor = fClsExtrato.Saldo
            fClsExtrato.ProcessamentoMensal
            .Row = sMes
            .RowHidden = False
            .SetText 1, sMes, MonthName(sMes)
            .SetText 2, sMes, Round(fClsExtrato.Saldo - sValor, 2)
            .SetText 3, sMes, Round(fClsExtrato.Saldo, 2)
        End With
    Next
Else
    MsgBox "Não é possível realizar uma aplicação futura, escolha um mês menor ou igual ao atual"
End If

Exit Sub
Resume
vbErrorHandler:
MsgBox Err.Number & " - " & Err.Description, vbOKOnly, Err.Source
End Sub

Private Function SomarValorMensal()
Dim sCont As Long, sValor As Double, sSaque As Double, sCont2 As Long
With gridAplicacoes
    For sCont = 1 To .MaxRows
        .Row = sCont
        .col = 1
        If .Text <> 0 Then
            .col = 2
            If .Text = MonthName(cmbMes.ListIndex) Then
                .col = 3
                sValor = CDbl(.Text)
                .col = 5
                sSaque = CDbl(.Text)
                fClsExtrato.Depositar sValor - sSaque
            ElseIf cmbMes.ListIndex > RetornaNumeroMes(.Text) Then
                For sCont2 = 1 To gridSimulacao.MaxRows
                    gridSimulacao.Row = sCont
                    gridSimulacao.col = 1
                    If gridSimulacao.Text <> "" Then
                        If gridSimulacao.Text = MonthName(cmbMes.ListIndex - 1) Then
                            gridSimulacao.col = 3
                            sValor = CDbl(gridSimulacao.Text)
                            fClsExtrato.Depositar sValor
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
        Else
            Exit For
        End If
    Next
End With
End Function

Private Function RetornaNumeroMes(ParMes As String) As Byte
Dim sRetorno As Byte


If ParMes = "janeiro" Then
    sRetorno = 1
ElseIf ParMes = "fevereiro" Then
    sRetorno = 2
ElseIf ParMes = "março" Then
    sRetorno = 3
ElseIf ParMes = "abril" Then
    sRetorno = 4
ElseIf ParMes = "maio" Then
    sRetorno = 5
ElseIf ParMes = "junho" Then
    sRetorno = 6
ElseIf ParMes = "julho" Then
    sRetorno = 7
ElseIf ParMes = "agosto" Then
    sRetorno = 8
ElseIf ParMes = "setembro" Then
    sRetorno = 9
ElseIf ParMes = "outubro" Then
    sRetorno = 10
ElseIf ParMes = "novembro" Then
    sRetorno = 11
ElseIf ParMes = "dezembro" Then
    sRetorno = 12
Else
    sRetorno = 0
End If

RetornaNumeroMes = sRetorno
End Function

Private Sub cmdOpcao_Click(Index As Integer)
On Error GoTo ErrorHandler

    If Index = EnumOption.Include Then
        fOpcao = Index
        DefinirTela True
        chkInativo.Visible = False
        quadDatas.Visible = False
        txtCodigo = "NOVO"
        chkInvestir.Enabled = False
    ElseIf Index = EnumOption.Read Or Index = EnumOption.Update Or Index = EnumOption.Delete Then
        gridPrincipal_Click gridPrincipal.ActiveCol, gridPrincipal.ActiveRow
        fOpcao = Index
        DefinirTela True
        If Not ObterDados Then GoTo ErrorHandler
        If fOpcao = EnumOption.Delete Then cmdOpcao_Click (EnumOption.Confirm)
    ElseIf Index = EnumOption.Confirm Then
        If fOpcao = EnumOption.Include Or fOpcao = EnumOption.Update Then
            TransferirDados
            DefinirTela False
            ExpurgarTela
        ElseIf fOpcao = EnumOption.Delete Then
            If Not fClsAplicacoes.Excluir(fCodigo) Then GoTo ErrorHandler
            AlimentarGrid
            DefinirTela False
            ExpurgarTela
        Else
            DefinirTela False
            ExpurgarTela
        End If
    ElseIf Index = EnumOption.Cancel Then
        DefinirTela False
        ExpurgarTela
    ElseIf Index = EnumOption.Leave Then
        Unload Me
    End If
Exit Sub

Exit Sub
Resume
ErrorHandler:
End Sub

Private Sub cmdLimparAplicacoes_Click()
If MsgBox("Deseja realmente limpar essa aplicação?", vbYesNo, "Aplicacções") = 6 Then
    SpreadClean gridAplicacoes
    SpreadClean gridSimulacao
    fClsExtrato.LimparExtrato
    fClsExtrato.LimparSaldos
    Set fClsExtrato = Nothing
End If
End Sub

Private Sub cmdPesquisa_Click()
AlimentarGrid
End Sub

Private Sub Form_Load()
cmdOpcao(EnumOption.Confirm).Visible = False
cmdOpcao(EnumOption.Cancel).Visible = False
quadCadastro.Visible = False
quadInvestimento.Visible = False
quadPesquisa.Visible = True
lblRodape = FillFooter
ComboBoxFill cmbCorretora, "SELECT COR_CODIGO, (COR_NOME + '(' + CONVERT(VARCHAR,COR_CODIGO) + ')') AS DESCRICAO FROM CORRETORAS WHERE COR_INATIVO = '0'"
cmbAno.AddItem Year(Date), 0
With cmbMes
    Dim sCont As Long
    .AddItem " ", 0
    For sCont = 1 To 12
        .AddItem MonthName(sCont) & "(" & sCont & ")", sCont
    Next
End With
AlimentarGrid
End Sub

Private Sub DefinirTela(ParCadastro As Boolean)
If ParCadastro = True Then
    quadPesquisa.Visible = False
    quadCadastro.Visible = True
    cmdOpcao(EnumOption.Include).Visible = False
    cmdOpcao(EnumOption.Read).Visible = False
    cmdOpcao(EnumOption.Update).Visible = False
    cmdOpcao(EnumOption.Delete).Visible = False
    cmdOpcao(EnumOption.Leave).Visible = False
    cmdOpcao(EnumOption.Confirm).Visible = True
    cmdOpcao(EnumOption.Cancel).Visible = True
Else
    quadPesquisa.Visible = True
    quadCadastro.Visible = False
    cmdOpcao(EnumOption.Include).Visible = True
    cmdOpcao(EnumOption.Read).Visible = True
    cmdOpcao(EnumOption.Update).Visible = True
    cmdOpcao(EnumOption.Delete).Visible = True
    cmdOpcao(EnumOption.Leave).Visible = True
    cmdOpcao(EnumOption.Confirm).Visible = False
    cmdOpcao(EnumOption.Cancel).Visible = False
End If

End Sub

Private Function ObterDados() As Boolean
If fClsAplicacoes.Consultar(fCodigo) Then
    With fClsAplicacoes
        txtCodigo.Text = .Codigo
        txtNome = .Nome
        cmbCorretora.ListIndex = .Corretora
        cmbAno.Text = .Ano
        chkInvestir = .Investir
        txtData = .Cadastro
        txtAtualizacao = .Atualizacao
        chkInativo = .Inativo
    End With
End If

fClsExtrato.CodigoAplicacao = CInt(txtCodigo)
fClsExtrato.ConsultarExtrato gridAplicacoes, txtValor, cmbMes, txtTaxa, txtSaque, cmdAplicar
If txtTaxa <> Empty Then
    txtTaxa.Enabled = False
Else
    txtTaxa.Enabled = True
    txtTaxa.Text = ""
End If
txtValor.Text = ""
cmbMes.Text = ""
txtValor.Text = ""
txtSaque.Text = ""
txtCodigo = fCodigo

ObterDados = True
End Function

Private Function TransferirDados() As Boolean
On Error GoTo Trata

Dim sSql As String, sCont As Long, sLinhas As Long, sMes As Byte, sRendimento As Double, sSaldo As Double

If fCodigo <> 0 Then fClsAplicacoes.Consultar (fCodigo)
With fClsAplicacoes
    .Codigo = fCodigo
    .Nome = txtNome
    .Corretora = cmbCorretora.ListIndex
    .Ano = cmbAno.Text
    .Investir = IIf(chkInvestir.Value, 1, 0)
    .Inativo = IIf(chkInativo.Value, 1, 0)
    If fOpcao = EnumOption.Include Then If Not .Inserir Then GoTo Trata
    If fOpcao = EnumOption.Update Then If Not .Atualizar Then GoTo Trata
End With

With fClsExtrato
    For sCont = 1 To gridAplicacoes.MaxRows
        sSql = "SELECT EXT_REGISTRO FROM EXTRATO WHERE EXT_REGISTRO = '" & CDbl(SpreadGetText(gridAplicacoes, 1, sCont)) & "'"
        ReadQuery sSql, sLinhas
        If sLinhas <> 0 Then GoTo Fim
        .Registro = CDbl(SpreadGetText(gridAplicacoes, 1, sCont))
        If .Registro = 0 Then Exit For
        .Valor = CDbl(SpreadGetText(gridAplicacoes, 3, sCont))
        .Taxa = txtTaxa
        .Saque = CDbl(SpreadGetText(gridAplicacoes, 5, sCont))
        .Mes = RetornaNumeroMes(SpreadGetText(gridAplicacoes, 2, sCont))
        If fOpcao = EnumOption.Update Then
            fClsExtrato.CodigoAplicacao = VariableAdjust(txtCodigo.Text, DoubleNumber)
            If Not .AtualizarExtrato() Then GoTo Trata
        End If
Fim:
    Next
    If fOpcao = EnumOption.Update Then
        If Not .LimparSaldos Then GoTo Trata
        For sCont = 1 To gridSimulacao.MaxRows
            gridSimulacao.Row = sCont
            gridSimulacao.col = 1
            If gridSimulacao.Text <> "0" Then
                gridSimulacao.col = 1
                sMes = RetornaNumeroMes(gridSimulacao.Text)
                gridSimulacao.col = 2
                sRendimento = CDbl(gridSimulacao.Text)
                gridSimulacao.col = 3
                sSaldo = CDbl(gridSimulacao.Text)
                If Not .InserirSaldos(sMes, sRendimento, sSaldo) Then GoTo Trata
            End If
        Next
    End If
End With

AlimentarGrid

TransferirDados = True
Exit Function
Resume
Trata:
MsgBox ErrorHandler(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Atualizar"
End Function

Private Function AnalisarDados() As Boolean

If CDbl(txtValor.Text) = 0 Then
    MsgBox "Por favor informe o valor da aplicação.", vbInformation
    txtValor.SetFocus
    Exit Function
End If

If Trim(cmbMes.Text) = "" Then
    MsgBox "Por favor informe o mês da aplicação.", vbInformation
    cmbMes.SetFocus
    Exit Function
End If

If CDbl(txtTaxa.Text) = 0 Then
    MsgBox "Por favor informe a taxa de investimento.", vbInformation
    txtTaxa.SetFocus
    Exit Function
End If

AnalisarDados = True
End Function

Private Sub ExpurgarTela()

txtCodigo = ""
txtNome = ""
chkInativo.Value = 0
txtData = ""
txtAtualizacao = ""
chkInativo.Visible = True
quadDatas.Visible = True
chkInvestir.Enabled = True
chkInvestir.Value = False
quadInvestimento.Visible = False
cmbCorretora.Text = ""
cmbAno.Text = ""
txtValor.Text = ""
cmbMes.Text = ""
txtTaxa.Text = ""
txtSaque.Text = ""
chkInvestir.Enabled = True
Set fClsExtrato = Nothing
Set fClsAplicacoes = Nothing

End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub gridPrincipal_Click(ByVal col As Long, ByVal Row As Long)
SpreadGetCode gridPrincipal, Row, fCodigo
End Sub

Private Sub gridPrincipal_DblClick(ByVal col As Long, ByVal Row As Long)
cmdOpcao_Click (EnumOption.Update)
End Sub
