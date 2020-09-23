VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpal 
   BackColor       =   &H00E9DEDB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SComboBox - fred_cpp & HACKPRO TM 2004 @ México - Colombia "
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   2850
      MouseIcon       =   "frmPpal.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtMaxListLenght 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   2985
      TabIndex        =   3
      Top             =   330
      Width           =   1590
   End
   Begin VB.TextBox txtAddItem 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   1530
      TabIndex        =   15
      Text            =   "HACKPRO TM"
      Top             =   1905
      Width           =   1245
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add Item"
      Height          =   405
      Left            =   240
      MouseIcon       =   "frmPpal.frx":0894
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1830
      Width           =   1110
   End
   Begin VB.TextBox txtSearchItem 
      ForeColor       =   &H00C56A31&
      Height          =   285
      Left            =   1530
      TabIndex        =   13
      Text            =   "fred_cpp"
      Top             =   1515
      Width           =   1245
   End
   Begin VB.CommandButton cmdTextItem 
      Caption         =   "Text Item"
      Height          =   405
      Left            =   1920
      MouseIcon       =   "frmPpal.frx":0B9E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   105
      Width           =   915
   End
   Begin VB.CommandButton cmdSearchItem 
      Caption         =   "Search Item"
      Height          =   405
      Left            =   240
      MouseIcon       =   "frmPpal.frx":0EA8
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1440
      Width           =   1110
   End
   Begin VB.CommandButton cmdDisabled 
      Caption         =   "&Enabled"
      Height          =   375
      Left            =   2850
      MouseIcon       =   "frmPpal.frx":11B2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   720
      Width           =   1125
   End
   Begin VB.ComboBox cmbAlign 
      ForeColor       =   &H00C56A31&
      Height          =   315
      ItemData        =   "frmPpal.frx":14BC
      Left            =   4770
      List            =   "frmPpal.frx":14C9
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   315
      Width           =   1860
   End
   Begin VB.CommandButton cmdListIndex 
      Caption         =   "ListIndex"
      Height          =   405
      Left            =   1020
      MouseIcon       =   "frmPpal.frx":14FD
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   105
      Width           =   915
   End
   Begin VB.Frame FramStyle 
      BackColor       =   &H00E9DEDB&
      Caption         =   "Combo Styles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   2700
      Left            =   120
      TabIndex        =   16
      Top             =   2430
      Width           =   6555
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   20
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "Office Xp"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   4
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighLightBorderColor=   8421504
         MaxListLength   =   -1
         NormalBorderColor=   8421504
         NumberItemsToShow=   -1
         SelectBorderColor=   4210752
         Text            =   "Office 2000"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   1095
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   7
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "Mac"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   12
         Left            =   120
         TabIndex        =   29
         Top             =   1455
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   10
         DisabledColor   =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighLightBorderColor=   12556415
         MaxListLength   =   -1
         MouseIcon       =   "frmPpal.frx":1807
         MousePointer    =   99
         NormalBorderColor=   12556415
         NumberItemsToShow=   -1
         SelectBorderColor=   12556415
         Style           =   1
         Text            =   "Picture"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   1
         Left            =   1710
         TabIndex        =   18
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   2
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         MouseIcon       =   "frmPpal.frx":1B21
         MousePointer    =   99
         NumberItemsToShow=   -1
         Text            =   "Win98"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   5
         Left            =   1710
         TabIndex        =   22
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   5
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NormalBorderColor=   14737632
         NumberItemsToShow=   -1
         Text            =   "Soft Style"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   9
         Left            =   1710
         TabIndex        =   26
         Top             =   1095
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   8
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NormalBorderColor=   12632256
         NumberItemsToShow=   -1
         Text            =   "JAVA"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   13
         Left            =   1710
         TabIndex        =   30
         Top             =   1455
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   11
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "Special Borde"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   2
         Left            =   3300
         TabIndex        =   19
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp Aqua"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   315
         Index           =   6
         Left            =   3300
         TabIndex        =   23
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   12
         ArrowColor      =   11491119
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighLightBorderColor=   10432512
         MaxListLength   =   -1
         NormalBorderColor=   13605023
         NumberItemsToShow=   -1
         SelectBorderColor=   10518399
         Text            =   "Circular"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   10
         Left            =   3300
         TabIndex        =   27
         Top             =   1095
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   9
         ArrowColor      =   4210752
         BackColor       =   16120314
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "Explorer Bar"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   11
         Left            =   4905
         TabIndex        =   28
         Top             =   1095
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp Silver"
         XpAppearance    =   3
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   16
         Left            =   120
         TabIndex        =   33
         Top             =   1830
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   13
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   16777215
         GradientColor2  =   14208709
         MaxListLength   =   -1
         NormalBorderColor=   12937777
         NumberItemsToShow=   -1
         SelectListBorderColor=   12937777
         Text            =   "GradientV"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   14
         Left            =   3300
         TabIndex        =   31
         Top             =   1455
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   14
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   16777215
         GradientColor2  =   14208709
         MaxListLength   =   -1
         NormalBorderColor=   12937777
         NumberItemsToShow=   -1
         SelectListBorderColor=   12937777
         Text            =   "GradientH"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   15
         Left            =   4905
         TabIndex        =   32
         Top             =   1455
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   15
         ArrowColor      =   0
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "Light Blue"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   3
         Left            =   4905
         TabIndex        =   20
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   6
         ArrowColor      =   33023
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   16776960
         GradientColor2  =   8421376
         HighLightBorderColor=   16777088
         HighLightColorText=   16744576
         MaxListLength   =   -1
         NormalBorderColor=   8421376
         NumberItemsToShow=   -1
         SelectBorderColor=   8421376
         Text            =   "KDE"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   7
         Left            =   4905
         TabIndex        =   24
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp Green"
         XpAppearance    =   2
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   0
         Left            =   1725
         TabIndex        =   34
         Top             =   1830
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp Gold"
         XpAppearance    =   5
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   17
         Left            =   3315
         TabIndex        =   35
         Top             =   1830
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp TasBlue"
         XpAppearance    =   4
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   18
         Left            =   4920
         TabIndex        =   36
         Top             =   1830
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         ArrowColor      =   4210752
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         Text            =   "WinXp Blue"
         XpAppearance    =   6
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   19
         Left            =   120
         TabIndex        =   37
         Top             =   2205
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   16
         ArrowColor      =   8675406
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   14332536
         GradientColor2  =   16768414
         HighLightBorderColor=   8675406
         MaxListLength   =   -1
         NormalBorderColor=   8675406
         NumberItemsToShow=   -1
         SelectBorderColor=   8675406
         SelectListBorderColor=   12937777
         Text            =   "Style Arrow"
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   21
         Left            =   1725
         TabIndex        =   38
         Top             =   2205
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   3
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighLightBorderColor=   255
         MaxListLength   =   -1
         NormalBorderColor=   16761024
         NumberItemsToShow=   -1
         Text            =   "WinXp Custom"
         XpAppearance    =   7
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   22
         Left            =   3315
         TabIndex        =   39
         Top             =   2205
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   17
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   16777215
         GradientColor2  =   12164479
         HighLightBorderColor=   8421504
         MaxListLength   =   -1
         NormalBorderColor=   8421504
         NumberItemsToShow=   -1
         SelectBorderColor=   8421504
         Text            =   "NiaWBSS"
         XpAppearance    =   4
      End
      Begin ComboBox.SComboBox XpComboBox2 
         Height          =   300
         Index           =   23
         Left            =   4920
         TabIndex        =   40
         Top             =   2205
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         AppearanceCombo =   18
         ArrowColor      =   8413007
         DisabledColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighLightBorderColor=   16771023
         MaxListLength   =   -1
         NormalBorderColor=   12576751
         NumberItemsToShow=   -1
         SelectBorderColor=   15775871
         Text            =   "Rhombus"
         XpAppearance    =   6
      End
   End
   Begin VB.CommandButton cmdListCount 
      Caption         =   "&ListCount"
      Height          =   405
      Left            =   120
      MouseIcon       =   "frmPpal.frx":1E3B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   105
      Width           =   915
   End
   Begin VB.ComboBox cmbStyle 
      ForeColor       =   &H00C56A31&
      Height          =   315
      ItemData        =   "frmPpal.frx":2145
      Left            =   240
      List            =   "frmPpal.frx":2185
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1860
   End
   Begin MSComctlLib.ImageList imgListIcon 
      Left            =   -765
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":22D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2870
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":381C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":422E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5574
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":624E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":67E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6942
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7354
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8104
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":97F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A124
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B804
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":BB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C138
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CB4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D0E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DAF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E508
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EF1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F2B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F652
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F9EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FD8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":101D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10622
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11306
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11752
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11FEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCombo 
      BackColor       =   &H00E9DEDB&
      Height          =   1590
      Left            =   4065
      ScaleHeight     =   1530
      ScaleWidth      =   2490
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   2550
      Begin VB.ComboBox ComboBox1 
         ForeColor       =   &H00C56A31&
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Text            =   "ComboBox1"
         Top             =   1125
         Width           =   2220
      End
      Begin ComboBox.SComboBox XpComboBox1 
         Height          =   315
         Left            =   165
         TabIndex        =   10
         Top             =   360
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         AppearanceCombo =   19
         ArrowColor      =   192
         AutoCompleteWord=   -1  'True
         BackColor       =   16381420
         DisabledColor   =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   14201464
         GradientColor2  =   16770226
         ListColor       =   16381420
         MaxListLength   =   -1
         MouseIcon       =   "frmPpal.frx":12306
         MousePointer    =   99
         NormalBorderColor=   8413007
         NumberItemsToShow=   -1
         Text            =   "HACKPRO TM"
         XpAppearance    =   7
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Combo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C56A31&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   43
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SComboBox Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C56A31&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   42
         Top             =   120
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "&RemoveItem"
      Height          =   375
      Left            =   2850
      MouseIcon       =   "frmPpal.frx":12620
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Image imgMaxListLenght 
      Height          =   240
      Left            =   4335
      MouseIcon       =   "frmPpal.frx":1292A
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":12C34
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaxListLength"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   4
      Left            =   2985
      TabIndex        =   45
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment Text List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   3
      Left            =   4785
      TabIndex        =   44
      Top             =   75
      Width           =   1635
   End
   Begin VB.Image imgAlign 
      Height          =   240
      Left            =   6390
      MouseIcon       =   "frmPpal.frx":12FBE
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":132C8
      Top             =   75
      Width           =   240
   End
   Begin VB.Image imgSetStyle 
      Height          =   240
      Left            =   1365
      MouseIcon       =   "frmPpal.frx":13652
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":1395C
      Top             =   615
      Width           =   240
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   195
      Index           =   0
      Left            =   255
      TabIndex        =   41
      Top             =   615
      Width           =   1065
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
          '***********************************'
          '* Copyright (C) 2004 - HACKPRO TM *'
          '*  Heriberto Mantilla Santamaría  *'
          '*        Barrancabermeja          *'
          '***********************************'
Option Explicit

 Private Const SW_SHOWMAXIMIZED = 3
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Private i As Integer

Private Sub cmdAddItem_Click()
 XpComboBox1.AddItem txtAddItem.Text, &HFE0099
End Sub

Private Sub cmdDisabled_Click()
 If (cmdDisabled.Caption = "&Disabled") Then
  XpComboBox1.Enabled = False
  cmdDisabled.Caption = "&Enabled"
 Else
  XpComboBox1.Enabled = True
  cmdDisabled.Caption = "&Disabled"
 End If
End Sub

Private Sub cmdHelp_Click()
 Call ShellExecute(frmPpal.hWnd, vbNullString, App.Path & "\Ayuda\Principal.html", vbNullString, "C:\", SW_SHOWMAXIMIZED)
End Sub

Private Sub cmdListCount_Click()
 MsgBox "ListCount: " & XpComboBox1.ListCount, vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub cmdListIndex_Click()
 MsgBox "ListIndex: " & XpComboBox1.ListIndex, vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub cmdRemoveItem_Click()
 XpComboBox1.RemoveItem 3
End Sub

Private Sub cmdSearchItem_Click()
 MsgBox "FindItem: " & XpComboBox1.FindItemText(txtSearchItem.Text, None), vbOKOnly + vbInformation, "SComboBox"
End Sub

Private Sub cmdTextItem_Click()
 MsgBox "ItemText: " & XpComboBox1.List(XpComboBox1.ListIndex), vbInformation + vbOKOnly, "SComboBox"
End Sub

Private Sub Form_Load()
 Me.Show
 For i = 1 To 18
  ComboBox1.AddItem "HACKPRO" & i
  If (i = 4) Or (i = 9) Or (i = 15) Or (i = 10) Then
   XpComboBox1.AddItem "HACKPRO" & i, , imgListIcon.ListImages(i).Picture, False
  ElseIf (i = 8) Or (i = 12) Then
   XpComboBox1.AddItem "HACKPRO" & i, &HFE0099, , , "Hola" & i
  ElseIf (i = 5) Or (i = 1) Or (i = 13) Then
   XpComboBox1.AddItem "HACKPRO" & i, &HFE0099, imgListIcon.ListImages(i).Picture, , , , , imgListIcon.ListImages(41).Picture, True
  Else
   XpComboBox1.AddItem "HACKPRO" & i, , imgListIcon.ListImages(i).Picture
  End If
 Next
 Set XpComboBox1.MouseIcon = imgListIcon.ListImages(41).Picture
 XpComboBox1.MousePointer = vbCustom
 Set XpComboBox1.NormalPictureUser = imgListIcon.ListImages(39).Picture
 Set XpComboBox1.DisabledPictureUser = imgListIcon.ListImages(40).Picture
 Set XpComboBox1.FocusPictureUser = imgListIcon.ListImages(39).Picture
 Set XpComboBox1.HighLightPictureUser = imgListIcon.ListImages(39).Picture
 For i = 1 To 3
  XpComboBox2(12).AddItem "Picture 0" & i
 Next
 XpComboBox2(12).ListIndex = 2
 Set XpComboBox2(12).MouseIcon = imgListIcon.ListImages(41).Picture
 XpComboBox2(12).MousePointer = vbCustom
 Call XpComboBox2_SelectionMade(12, "Picture 02", 2)
 XpComboBox1.MaxListLength = 19
 XpComboBox1.NumberItemsToShow = 8
 XpComboBox1.Text = XpComboBox1.List(2)
 cmbStyle.ListIndex = XpComboBox1.AppearanceCombo - 1
 cmbAlign.ListIndex = XpComboBox1.Alignment
 txtMaxListLenght.Text = XpComboBox1.MaxListLength
 imgSetStyle_Click
 If (XpComboBox1.Enabled = True) Then cmdDisabled.Caption = "&Disabled"
End Sub

Private Sub imgAlign_Click()
 XpComboBox1.Alignment = cmbAlign.ListIndex
End Sub

Private Sub imgMaxListLenght_Click()
 Dim isValue As Long
 
 isValue = CLng(txtMaxListLenght.Text)
 If (isValue > 0) And (IsNumeric(isValue) = True) Then XpComboBox1.MaxListLength = isValue
End Sub

Private Sub imgSetStyle_Click()
 XpComboBox1.AppearanceCombo = cmbStyle.ListIndex + 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 8) Then Exit Sub
 If (IsNumeric(Chr$(KeyAscii)) = False) Then KeyAscii = 0: Beep
End Sub

Private Sub XpComboBox2_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
 If (Index = 12) Then
  Select Case SelectedItem
   Case "Picture 01"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(32).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(29).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(31).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(30).Picture
    XpComboBox2(12).NormalBorderColor = &HB99D7F
    XpComboBox2(12).SelectBorderColor = &HC56A31
    XpComboBox2(12).HighLightBorderColor = &HC56A31
   Case "Picture 02"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(36).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(34).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(35).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(33).Picture
    XpComboBox2(12).NormalBorderColor = &H9F989F
    XpComboBox2(12).SelectBorderColor = &H406790
    XpComboBox2(12).HighLightBorderColor = &H90887F
   Case "Picture 03"
    Set XpComboBox2(12).FocusPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).HighLightPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).NormalPictureUser = imgListIcon.ListImages(37).Picture
    Set XpComboBox2(12).DisabledPictureUser = imgListIcon.ListImages(38).Picture
    XpComboBox2(12).NormalBorderColor = &H103030
    XpComboBox2(12).SelectBorderColor = &H103030
    XpComboBox2(12).HighLightBorderColor = &H103030
  End Select
 End If
End Sub
