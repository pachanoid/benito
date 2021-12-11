VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmJoystick 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15780
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   692
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1052
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   9360
      TabIndex        =   91
      Text            =   "Text12"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.ComboBox cmbJoy 
      BackColor       =   &H80000007&
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4106
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   90
      Top             =   8760
      Width           =   2775
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   8880
      TabIndex        =   71
      Text            =   "Text11"
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "OFF"
      Height          =   615
      Left            =   9600
      TabIndex        =   70
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ON"
      Height          =   615
      Left            =   8880
      TabIndex        =   69
      Top             =   9120
      Width           =   735
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Control 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4080
      TabIndex        =   66
      Top             =   6120
      Width           =   2055
      Begin VB.PictureBox Joy 
         BackColor       =   &H00000000&
         Height          =   1380
         Index           =   2
         Left            =   360
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   67
         Top             =   480
         Width           =   1380
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000016&
      Caption         =   "Z Axis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   64
      Top             =   8160
      Width           =   2295
      Begin VB.PictureBox Joy 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1500
         Index           =   3
         Left            =   360
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   65
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   4080
      TabIndex        =   58
      Text            =   "Text10"
      Top             =   10080
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   4080
      TabIndex        =   57
      Text            =   "Text9"
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   4080
      TabIndex        =   56
      Text            =   "Text4"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   4080
      TabIndex        =   55
      Text            =   "Text3"
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   52
      Text            =   "Text3"
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   51
      Text            =   "Text3"
      Top             =   9840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000C000&
      Height          =   1575
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      Top             =   7200
      Width           =   2775
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   8160
      TabIndex        =   47
      Text            =   "23"
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   6840
      TabIndex        =   46
      Text            =   "192.168.1.177"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000015&
      Caption         =   "conectar"
      Height          =   495
      Left            =   6120
      MaskColor       =   &H80000015&
      TabIndex        =   45
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Timer tmrScan 
      Interval        =   10
      Left            =   15000
      Top             =   9360
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Control 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2280
      TabIndex        =   33
      Top             =   6120
      Width           =   1815
      Begin VB.PictureBox Joy 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1380
         Index           =   1
         Left            =   240
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   34
         Top             =   480
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Botonera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   26
      Top             =   9120
      Width           =   2775
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   195
         Index           =   19
         Left            =   2235
         TabIndex        =   44
         Top             =   840
         Width           =   195
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   43
         Top             =   840
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         Height          =   195
         Index           =   17
         Left            =   1320
         TabIndex        =   42
         Top             =   840
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         Height          =   195
         Index           =   16
         Left            =   840
         TabIndex        =   41
         Top             =   840
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   40
         Top             =   840
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   39
         Top             =   480
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   38
         Top             =   480
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   195
         Index           =   12
         Left            =   1320
         TabIndex        =   37
         Top             =   480
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   36
         Top             =   480
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   105
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   5
         Left            =   240
         Shape           =   5  'Rounded Square
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   8
         Left            =   1680
         Shape           =   5  'Rounded Square
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   7
         Left            =   1200
         Shape           =   5  'Rounded Square
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   6
         Left            =   720
         Shape           =   5  'Rounded Square
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   4
         Left            =   2160
         Shape           =   5  'Rounded Square
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   3
         Left            =   1680
         Shape           =   5  'Rounded Square
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   2
         Left            =   1200
         Shape           =   5  'Rounded Square
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   1
         Left            =   720
         Shape           =   5  'Rounded Square
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   0
         Left            =   240
         Shape           =   5  'Rounded Square
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Butt 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   9
         Left            =   2160
         Shape           =   5  'Rounded Square
         Top             =   720
         Width           =   375
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         Height          =   195
         Index           =   6
         Left            =   1320
         TabIndex        =   32
         Top             =   930
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   31
         Top             =   930
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   30
         Top             =   450
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   29
         Top             =   450
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   28
         Top             =   450
         Width           =   105
      End
      Begin VB.Label NoX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   450
         Width           =   105
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "Control PT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   16
      Top             =   6120
      Width           =   2295
      Begin VB.PictureBox Joy 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawWidth       =   2
         FillColor       =   &H000000C0&
         FillStyle       =   3  'Vertical Line
         ForeColor       =   &H000000C0&
         Height          =   1500
         Index           =   0
         Left            =   360
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   17
         Top             =   360
         Width           =   1500
         Begin VB.PictureBox Arw 
            BackColor       =   &H80000012&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   480
            Picture         =   "panel.frx":0000
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   480
            Picture         =   "panel.frx":07AE
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BackColor       =   &H80000012&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   960
            Picture         =   "panel.frx":0F5C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   0
            Picture         =   "panel.frx":170A
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   720
            Picture         =   "panel.frx":1EB8
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   5
            Left            =   240
            Picture         =   "panel.frx":2666
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   20
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BackColor       =   &H80000012&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   7
            Left            =   240
            Picture         =   "panel.frx":2E14
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Arw 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   720
            Picture         =   "panel.frx":35C2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      TabIndex        =   5
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ".dwButtons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   840
         TabIndex        =   15
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ".dwPOV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ".dwZpos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   1155
         TabIndex        =   13
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ".dwXpos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1140
         TabIndex        =   12
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ".dwYpos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1140
         TabIndex        =   11
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblInfoEX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   10
         Left            =   2160
         TabIndex        =   10
         ToolTipText     =   "Buttons Estatus"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblInfoEX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   8
         Left            =   2160
         TabIndex        =   9
         ToolTipText     =   "Buttons Estatus"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblInfoEX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   4
         Left            =   2160
         TabIndex        =   8
         ToolTipText     =   "Buttons Estatus"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblInfoEX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   3
         Left            =   2160
         TabIndex        =   7
         ToolTipText     =   "Buttons Estatus"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblInfoEX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Buttons Estatus"
         Top             =   1080
         Width           =   1575
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6135
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   9855
      ExtentX         =   17383
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000015&
      Caption         =   "Desconectar"
      Height          =   495
      Left            =   9000
      MaskColor       =   &H00404040&
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser ControlWeb 
      CausesValidation=   0   'False
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      ExtentX         =   11245
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   9000
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   15000
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000016&
      Caption         =   "Valores Game Pad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2280
      TabIndex        =   73
      Top             =   8160
      Width           =   3855
      Begin VB.Label Label24 
         BackColor       =   &H80000016&
         Caption         =   "a"
         Height          =   255
         Left            =   1440
         TabIndex        =   89
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000016&
         Caption         =   "b"
         Height          =   255
         Left            =   1440
         TabIndex        =   88
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000016&
         Caption         =   "y1"
         Height          =   255
         Left            =   1320
         TabIndex        =   87
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000016&
         Caption         =   "y1"
         Height          =   255
         Left            =   3360
         TabIndex        =   86
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000016&
         Caption         =   "b"
         Height          =   255
         Left            =   3360
         TabIndex        =   85
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000016&
         Caption         =   "a"
         Height          =   255
         Left            =   3360
         TabIndex        =   84
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000016&
         Caption         =   "x1"
         Height          =   375
         Left            =   3360
         TabIndex        =   83
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Caracter enviado"
         Height          =   255
         Left            =   1800
         TabIndex        =   82
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
         Caption         =   "Caracter enviado"
         Height          =   255
         Left            =   0
         TabIndex        =   81
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "x1"
         Height          =   375
         Left            =   1320
         TabIndex        =   80
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000016&
         Caption         =   "b"
         Height          =   255
         Left            =   1320
         TabIndex        =   79
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "y1"
         Height          =   255
         Left            =   2160
         TabIndex        =   78
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000016&
         Caption         =   "a"
         Height          =   255
         Left            =   1320
         TabIndex        =   77
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   0
         TabIndex        =   76
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   0
         TabIndex        =   75
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label14 
         Height          =   255
         Left            =   0
         TabIndex        =   74
         Top             =   1680
         Width           =   255
      End
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      Caption         =   "Caracter enviado"
      Height          =   255
      Left            =   4080
      TabIndex        =   72
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblInfoEX 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   1
      Left            =   10560
      TabIndex        =   68
      ToolTipText     =   "Buttons Estatus"
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "d"
      Height          =   375
      Left            =   5520
      TabIndex        =   63
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "c"
      Height          =   255
      Left            =   5520
      TabIndex        =   62
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "y1"
      Height          =   375
      Left            =   5520
      TabIndex        =   61
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "z1"
      Height          =   255
      Left            =   5520
      TabIndex        =   60
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label lblInfoEX 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   0
      Left            =   10560
      TabIndex        =   59
      ToolTipText     =   "Buttons Estatus"
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000016&
      Caption         =   "Port"
      Height          =   255
      Left            =   6120
      TabIndex        =   54
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000016&
      Caption         =   "ip"
      Height          =   255
      Left            =   6120
      TabIndex        =   53
      Top             =   6240
      Width           =   735
   End
End
Attribute VB_Name = "frmJoystick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x1, y1, z1, R1, u1, a, b, c, d, ab, abin, bbin, fa, fb, boton, letraA, letraB, letraC, letraD, letraE, letraF, letraG, letraH, valor1, valor2, valor3, valor4 As String
Dim anterior, anterior2, anterior3, anterior4, n As String
Dim valor As Integer



Private Sub Command2_Click()
'Actuador 1
z1 = Chr(9) & Chr(40)
Winsock1.SendData z1
Text11.Text = z1
End Sub

Private Sub Command4_Click()
'Actuador 1
z1 = Chr(10) & Chr(120)
Winsock1.SendData z1
Text11.Text = z1
End Sub
Private Sub Command3_Click()
On Error GoTo ErrSub

    With Winsock1
        .Close
        .RemoteHost = txtIP
        .RemotePort = txtPort
        .Connect
    End With
Exit Sub
ErrSub:
MsgBox "Error : " & Err.Description, vbCritical
End Sub




Private Sub Form_Load()

n = 0

'Init values
xZbk = -2000
xPOV = JOY_POVCENTERED
cmbJoy.Tag = "-1"
DrawMk  'Go to DrawMk Sub...
DrawOd  'Go to DrawOd Sub...
InitJoy 'Go to InitJoy Sub...

'Manejo de Observer Cam -------------

' Elimina la barra de direcciones
    ControlWeb.AddressBar = False
       
    'Elimina la barra de menú
    ControlWeb.MenuBar = False
       
    'Elimina la barra de herramientas
    ControlWeb.ToolBar = False
       
    'Elimina la barra de herramientas
    ControlWeb.StatusBar = False
    
    
     ' Navega a la página indicad **** activar con el ip adecuado para ver camaras
    ControlWeb.Navigate "http://192.168.1.134"

    'OBS Cam
    WebBrowser1.Navigate "http://192.168.1.148"

 '------------------------------------------



'If Not StartJoystick Then
 '   MsgBox "Conecte el joystick!"
'    End
'End If
'
' tcpClient.RemoteHost = InputBox("Direccion IP Observer2", _
'        "Direccion IP", "localhost")
'
'    If tcpClient.RemoteHost = "" Then
'        tcpClient.RemoteHost = "localhost"
'    End If
    
    
  
End Sub


Function Binary(InptD As Variant)

'Dim Binary As String
Dim BaseD, G As Integer

BaseD = 2
Binary = ""
G = InptD
Do
Binary = (G Mod BaseD) & Binary
G = G \ BaseD
Loop Until G = 0
End Function
Private Sub Command1_Click()
 On Local Error Resume Next
    If Not ControlWeb Is Nothing Then
        'cierra la ventana del navegador si estaba abierta
        ControlWeb.Quit
        Set ControlWeb = Nothing
    End If
Call tcpClient.Close
End Sub

Private Sub Form_Resize()
If frmJoystick.WindowState = 0 Then
    Sound MPathX(App.Path) & "snd\Up.wav", _
            SND_ASYNC + SND_NODEFAULT
Else
    Sound MPathX(App.Path) & "snd\Dn.wav", _
            SND_ASYNC + SND_NODEFAULT
End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'Exit App

End Sub


Private Sub mnuExit_Click()
On Error Resume Next
Kill (MPathX(App.Path) & "Joy.bat") 'Erase the .bat file
End

End Sub


Private Sub tmrScan_Timer()
Dim i As Long

fa = 0
fb = 0
ab = ""

On Error GoTo error_ciclo
'The scan joystick sub
With InfoJoyEX
    DisButtX cmbJoy.ListIndex
    
    IDJoy = joyGetPosEx(cmbJoy.ListIndex, InfoJoyEX) 'Get joystick info.
    'stbInf.Panels(1).Text = " Joystick Status: " & JoyEst(IDJoy) 'Update the status bar
    
    If .dwPOV <> Val(lblInfoEX(10).Caption) Then
        bPOV = True
    End If
            
    '*** Update labels with InfoJoyEX ***
'    lblInfoEX(13).Caption = Str(IDJoy)
'    lblInfoEX(0).Caption = .dwSize
'    lblInfoEX(1).Caption = .dwFlags
    lblInfoEX(2).Caption = .dwXpos
    lblInfoEX(3).Caption = .dwYpos
 '   -------------------------------------





x1 = (Round(.dwXpos / 350, 0))       'stick 2 eje x
y1 = (Round((.dwYpos / 350), 0))     'stick 2 eje Y

R1 = (Round(.dwRpos / 350, 0))         'stick 1 eje R y
u1 = (Round(.dwUpos / 350, 0))         'stick 1 eje U x


a = (x1 + y1) - 90
If a > 160 Then         'limita valores maximos y minimos entendidos por arduino
a = 160
End If

If a < 30 Then
a = 30
End If

If (a = 91 Or a = 92) Then
a = 90
End If
If (a = 92 Or a = 93) Then
a = 90
End If

If (a = 93 Or a = 94) Then
a = 90
End If
If (a = 95 Or a = 96) Then
a = 90
End If

If (a = 89 Or a = 88) Then
a = 90
End If
If (a = 87 Or a = 86) Then
a = 90
End If

If (a = 86 Or a = 85) Then
a = 90
End If
If (a = 84 Or a = 83) Then
a = 90
End If



b = (y1 - x1) + 90
If b > 160 Then
b = 160
End If


If b < 30 Then
b = 30
End If

If (b = 91 Or b = 92) Then
b = 90
End If
If (b = 92 Or b = 93) Then
b = 90
End If

If (b = 93 Or b = 94) Then
b = 90
End If
If (b = 95 Or b = 96) Then
b = 90
End If

If (b = 89 Or b = 88) Then
b = 90
End If
If (b = 87 Or b = 86) Then
b = 90
End If

If (b = 86 Or b = 85) Then
b = 90
End If
If (b = 84 Or b = 83) Then
b = 90
End If






Text1.Text = x1
Text2.Text = y1
Text3.Text = u1
Text4.Text = R1

Text5.Text = a
Text6.Text = b
Text9.Text = c
Text10.Text = d


letraA = Chr(1) 'motor1
letraB = Chr(a) ' pasa el valor de x1 a un caracter asii letraA se define como string

letraC = Chr(2) 'motor2
letraD = Chr(b)

letraE = Chr(3) ' servo1
letraF = Chr(u1)

letraG = Chr(4) ' servo2
letraH = Chr(R1)



valor1 = letraA & letraB
Text7.Text = valor1

valor2 = letraC & letraD
Text8.Text = valor2

valor3 = letraE & letraF       'servo1
Text9.Text = valor3

valor4 = letraG & letraH        'servo2
Text10.Text = valor4

'******* envia los valores a los motores y servos solo si cambian de estado
If valor1 <> anterior Then
Winsock1.SendData valor1

anterior = valor1
End If

If valor2 <> anterior2 Then
Winsock1.SendData valor2

anterior2 = valor2
End If

If valor3 <> anterior3 Then
Winsock1.SendData valor3

anterior3 = valor3
End If

If valor4 <> anterior4 Then
Winsock1.SendData valor4

anterior4 = valor4
End If


boton = .dwButtons
Text12.Text = boton

  If boton = 1 Then
  z1 = Chr(9) & Chr(40)
Winsock1.SendData z1
Text11.Text = z1

Else
z1 = Chr(10) & Chr(40)
Winsock1.SendData z1
Text11.Text = z1
    End If
 '----------------------------------------------
    lblInfoEX(4).Caption = .dwZpos
    lblInfoEX(0).Caption = .dwRpos  'stick1 eje y
    lblInfoEX(1).Caption = .dwUpos  'stick1 eje x
'    lblInfoEX(7).Caption = .dwVpos
    lblInfoEX(8).Caption = .dwButtons
'    lblInfoEX(9).Caption = .dwButtonNumber
    lblInfoEX(10).Caption = .dwPOV
    
    If .dwPOV = 0 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('U,5')") 'up
    If .dwPOV = 18000 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('D,5')") 'down
    If .dwPOV = 27000 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('L,5')") 'left
    If .dwPOV = 9000 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('R,5')") 'rigth
    If .dwPOV = 31500 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('UL,5')") 'up left
    If .dwPOV = 4500 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('UR,5')")  'up rigth
    If .dwPOV = 22500 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('DL,5')")
    If .dwPOV = 15500 Then ControlWeb.Navigate2 ("http://192.168.2.4:1024/pt/ptctrl.cgi?mv=('DR,5')")
    
    
        
    
'    lblInfoEX(11).Caption = .dwReserved1
'    lblInfoEX(12).Caption = .dwReserved2
            
    '*** Update graphical controls ***
    ButtonsX .dwButtons
    DrawAr .dwZpos
    DrawCz 1, .dwXpos, .dwYpos
    DrawCz 2, .dwRpos, .dwUpos
    DrawCd .dwPOV
    
    
End With

Exit Sub

error_ciclo:

Open App.Path & "\observer.log" For Append Access Read Write As #9
            Print #9, Date & " " & Time() & " observer--> " & Err.Description
 Close #9
 
End Sub

'****************************************
'<<<<  Personalized Sub-procedures.  >>>>
'****************************************

Public Sub ButtonsX(ByVal xBs As Long)
'Init buttons scan procedure...
'For 1 to 32 buttons.
SetButtX xBs, 0  'Go to SetButX Sub...
SetButtX xBs, 1
SetButtX xBs, 2
SetButtX xBs, 3
SetButtX xBs, 4
SetButtX xBs, 5
SetButtX xBs, 6
SetButtX xBs, 7
SetButtX xBs, 8
SetButtX xBs, 9

End Sub

Public Sub DisButtX(ByVal xNbu As Integer)
'Set Enebled o Disabled the joystick buttons
Dim xRj As Long
Dim xCo As Integer
If (xNbu >= 0) And (xNbu <> Val(cmbJoy.Tag)) Then
    xRj = joyGetDevCaps(cmbJoy.ListIndex, CapX, Len(CapX))  'With joyGetDevCaps function
    For xCo = 0 To 9
        If xCo < CapX.wNumButtons Then
            Butt(xCo).BackColor = G01
            Butt(xCo).BorderColor = BLK
'            NoX(xCo).ForeColor = BLK
        Else
'            Butt(xCo).BackColor = G02
'            Butt(xCo).BorderColor = G01
'            NoX(xCo).ForeColor = G01
        End If
    Next xCo
    cmbJoy.Tag = Trim(Str(xNbu))
End If

End Sub

Public Sub DrawAr(ByVal dwRes As Long)
'Move de arrow of the Z Axis and play sound (scroll left or scroll right)
Dim Med, R1, R2, Ang As Integer
If dwRes <> xZbk Then
    Med = Int(Joy(3).ScaleWidth / 2)
    R1 = Joy(3).ScaleWidth - 25
    R2 = Joy(3).ScaleWidth - 33

    Joy(3).Cls
    DrawOd
    Joy(3).Scale (-1 * Med, R1 + 15)-(Med, -15)

    Ang = Int(50 + dwRes * 80 / 65535)

    Joy(3).DrawWidth = 2

    Joy(3).Line (0, 0)- _
            (R1 * Cos(Ang * PI / 180), R1 * Sin(Ang * PI / 180)), GR
    Ang = Ang + 180
    Joy(3).Line (0, 0)- _
            (10 * Cos(Ang * PI / 180), 10 * Sin(Ang * PI / 180)), GR
                
    Joy(3).Circle (0, 0), 5, BLK
    
    If ((dwRes - SenX) > xZbk) Then
        If InSn Then
            InSn = False
            Sound MPathX(App.Path) & "snd\Dn.wav", _
                 SND_ASYNC + SND_NODEFAULT
        End If
    End If
    If ((dwRes + SenX) < xZbk) Then
        If InSn Then
            InSn = False
            Sound MPathX(App.Path) & "snd\Up.wav", _
                 SND_ASYNC + SND_NODEFAULT
        End If
    End If
    Joy(3).DrawWidth = 1
Else
    InSn = True
End If
xZbk = dwRes

End Sub

Public Sub DrawCd(ByVal dwPO As Long)
'Select what direction in POV is pressed
Dim xCa As Integer
If dwPO <> xPOV Then
 For xCa = 0 To 7
    If Arw(xCa).Visible Then
        Arw(xCa).Visible = False
    End If
 Next xCa

 Select Case dwPO
    Case JOY_POVFORWARD   'Dec 0
        DrawPOV 0
    Case JOY_POVFRDRHT 'Dec 4500
        DrawPOV 1
    Case JOY_POVRIGHT 'Dec 9000
        DrawPOV 2
    Case JOY_POVBRDRHT 'Dec 13500
        DrawPOV 3
    Case JOY_POVBACKWARD 'Dec 18000
        DrawPOV 4
    Case JOY_POVBRDLFT 'Dec 22500
        DrawPOV 5
    Case JOY_POVLEFT 'Dec 27000
        DrawPOV 6
    Case JOY_POVFRDLFT 'Dec 31500
        DrawPOV 7
 End Select
End If
xPOV = dwPO

End Sub

Public Sub DrawCz(ByVal xInd As Integer, ByVal Xval As Long, ByVal Yval As Long)
'Show de "+" pointer position for the X+Y Axis an U+R Axis.
Dim Xwi, Yhe As Long
Xwi = Int(Xval * (Joy(xInd).ScaleWidth - 1) / 65535)
Yhe = Int(Yval * (Joy(xInd).ScaleHeight - 1) / 65535)

Joy(xInd).Cls

Joy(xInd).Line (Xwi - 10, Yhe)- _
            (Xwi + 10, Yhe), GR
Joy(xInd).Line (Xwi, Yhe - 10)- _
            (Xwi, Yhe + 10), GR

If Xval < XY4D Then
    If bXYs(xInd, 0) Then
        Sound MPathX(App.Path) & "snd\Ax.wav", _
                SND_ASYNC + SND_NODEFAULT
        bXYs(xInd, 0) = False
    End If
Else
    bXYs(xInd, 0) = True
End If

If Xval > XY4U Then
    If bXYs(xInd, 1) Then
        Sound MPathX(App.Path) & "snd\Ax.wav", _
                SND_ASYNC + SND_NODEFAULT
        bXYs(xInd, 1) = False
    End If
Else
    bXYs(xInd, 1) = True
End If

If Yval < XY4D Then
    If bXYs(xInd, 2) Then
        Sound MPathX(App.Path) & "snd\Ax.wav", _
                SND_ASYNC + SND_NODEFAULT
        bXYs(xInd, 2) = False
    End If
Else
    bXYs(xInd, 2) = True
End If

If Yval > XY4U Then
    If bXYs(xInd, 3) Then
        Sound MPathX(App.Path) & "snd\Ax.wav", _
                SND_ASYNC + SND_NODEFAULT
        bXYs(xInd, 3) = False
    End If
Else
    bXYs(xInd, 3) = True
End If

End Sub

Public Sub DrawMk()
'Draw the POV-Graph marks ans set the arrow positions.
Dim Med, R1, R2, Ang As Integer
Med = Int(Joy(0).ScaleWidth / 2)
R1 = Med - 10
R2 = Med - 2
Joy(0).Scale (-1 * Med, Med)-(Med, -1 * Med)

For Ang = 0 To 360 Step 45
    Joy(0).Line (R1 * Cos(Ang * PI / 180), R1 * Sin(Ang * PI / 180))- _
                (R2 * Cos(Ang * PI / 180), R2 * Sin(Ang * PI / 180)), BLK
Next Ang
Arw(0).Move R1 * Cos(90 * PI / 180) - 1 * Arw(0).Width / 2, R1 * Sin(90 * PI / 180)
Arw(1).Move R1 * Cos(45 * PI / 180) - 1 * Arw(0).Width, R1 * Sin(45 * PI / 180)
Arw(2).Move R1 * Cos(0) - 1 * Arw(0).Width, R1 * Sin(0) + 1 * Arw(0).Width / 2
Arw(3).Move R1 * Cos(315 * PI / 180) - 1 * Arw(0).Width, R1 * Sin(315 * PI / 180) + Arw(0).Width
Arw(4).Move R1 * Cos(270 * PI / 180) - 1 * Arw(0).Width / 2, R1 * Sin(270 * PI / 180) + Arw(0).Width
Arw(5).Move R1 * Cos(225 * PI / 180), R1 * Sin(225 * PI / 180) + Arw(0).Width
Arw(6).Move R1 * Cos(180 * PI / 180), Arw(0).Width / 2
Arw(7).Move R1 * Cos(135 * PI / 180), R1 * Sin(135 * PI / 180)

Joy(0).AutoRedraw = False

End Sub

Public Sub DrawOd()
'Draw the Z-Graph marks.
Dim Med, R1, R2, Ang As Integer
'Med = Int(Joy(3).ScaleWidth / 2)
'R1 = Joy(3).ScaleWidth - 25
'R2 = Joy(3).ScaleWidth - 33
'Joy(3).Scale (-1 * Med, R1 + 15)-(Med, -15)

For Ang = 50 To 130 Step 10
  '  Joy(3).Line (R1 * Cos(Ang * PI / 180), R1 * Sin(Ang * PI / 180))- _
                (R2 * Cos(Ang * PI / 180), R2 * Sin(Ang * PI / 180)), _
                IIf(Ang < 80, R01, BLK)
Next Ang
'Joy(3).AutoRedraw = False

End Sub

Public Sub DrawPOV(ByVal xInd As Integer)
'Show the arrow in the POV
If Not (Arw(xInd).Visible) Then
    Arw(xInd).Visible = True
End If
If bPOV Then
    bPOV = False
    Sound MPathX(App.Path) & "snd\XY.wav", _
            SND_ASYNC + SND_NODEFAULT
End If

End Sub

Public Sub InitJoy()
'Get the joystick number in the system and about information.
Dim xJa, xRj As Long
Dim xJn As Integer
Dim sDJ As String
xJa = joyGetNumDevs()
For xJn = 0 To xJa
    xRj = joyGetDevCaps(xJn, CapX, Len(CapX))
    If Val(CapX.wPid) <> 0 Then
        sDJ = "Joistick " & Trim(Str(xJn + 1))
        sDJ = sDJ & Str(CapX.wNumAxes) & " Axes"
        sDJ = sDJ & Str(CapX.wNumButtons) & " Buttons"
        sDJ = sDJ & " (Joystick Microsoft Controler)"
        cmbJoy.AddItem sDJ
    End If
Next xJn
If cmbJoy.ListCount > 0 Then
    cmbJoy.ListIndex = 0
End If

InfoJoyEX.dwFlags = JOY_RETURNALL 'Dec 128: Request for status buttons.
InfoJoyEX.dwSize = &H40 'Dec 64
'DisButtX cmbJoy.ListIndex

End Sub

Public Sub SetButtX(ByVal xBi As Integer, ByVal nInd As Integer)
'Set the status button and play sound only one time.
If Val(Mid(DecBin(xBi, FSIZE), FSIZE - nInd, 1)) = 1 Then
    If Butt(nInd).BackColor = G01 Then
        Butt(nInd).BackColor = R01
    End If
    If BnSn(nInd) Then '
        BnSn(nInd) = False
        Sound MPathX(App.Path) & "snd\s" & IIf(nInd < 9, Trim(Str(nInd)), "X") & ".wav", _
                SND_ASYNC + SND_NODEFAULT
    End If
Else
    If Butt(nInd).BackColor = R01 Then
        Butt(nInd).BackColor = G01
    End If
    BnSn(nInd) = True
End If

If Butt(1) = True Then  'tratando de q el boton 1 haga algo
     Beep
     End If
     


End Sub

Private Sub Winsock1_Close()

    Winsock1.Close  'Cierra la conexión
    txtLog = txtLog & "*** Desconectado" & vbCrLf

End Sub

Private Sub Winsock1_Connect()

txtLog = "Conectado a " & Winsock1.RemoteHostIP & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim dat As String
    
    Winsock1.GetData dat, vbString
    txtLog = txtLog & "Servidor : " & dat & vbCrLf

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, _
                           Description As String, _
                           ByVal Scode As Long, _
                           ByVal Source As String, _
                           ByVal HelpFile As String, _
                           ByVal HelpContext As Long, _
                           CancelDisplay As Boolean)

    txtLog = txtLog & "*** Error : " & Description & vbCrLf

    Winsock1_Close
End Sub


