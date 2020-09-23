VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000040&
   Caption         =   "555 Timer Calculator"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   8160
      ScaleHeight     =   735
      ScaleWidth      =   660
      TabIndex        =   60
      Top             =   105
      Width           =   690
      Begin VB.Image Image5 
         Height          =   825
         Left            =   -45
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   810
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   180
      Top             =   45
   End
   Begin Project1.sTab sTab1 
      Height          =   555
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   64
      BackColorNormal =   16777215
      BackColorSelected=   12640511
      BorderColor     =   4210752
      BorderShadowColor=   -2147483633
      BorderShadowColorSelected=   -2147483632
      SpacingTop      =   3
      SpacingTopSelected=   0
      SpacingDown     =   3
      SpacingSides    =   10
      ListCount       =   4
      List1           =   "Tab 1"
      List2           =   "Tab 2"
      List3           =   "Tab 3"
      List4           =   "Tab 4"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C000C0&
      Height          =   6060
      Left            =   150
      TabIndex        =   1
      Top             =   1005
      Width           =   8700
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5595
         Left            =   6195
         TabIndex        =   14
         Top             =   270
         Width           =   2355
         Begin VB.PictureBox picAS_output 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   945
            Left            =   120
            ScaleHeight     =   61
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   140
            TabIndex        =   46
            Top             =   3660
            Width           =   2130
            Begin VB.PictureBox pic1 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   555
               Left            =   -1530
               ScaleHeight     =   37
               ScaleMode       =   0  'User
               ScaleWidth      =   306
               TabIndex        =   57
               Top             =   180
               Width           =   4620
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Vcc"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   0
               TabIndex        =   48
               Top             =   -45
               Width           =   300
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   15
               TabIndex        =   47
               Top             =   735
               Width           =   90
            End
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Animate"
            Height          =   240
            Left            =   105
            TabIndex        =   58
            Top             =   4635
            Width           =   1125
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Not actual time"
            Height          =   180
            Left            =   615
            TabIndex        =   59
            Top             =   3435
            Width           =   1125
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Off = .693 * R2 * C"
            Height          =   210
            Left            =   930
            TabIndex        =   44
            Top             =   2265
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "On = .693 * (R1 + R2) * C"
            Height          =   240
            Left            =   480
            TabIndex        =   43
            Top             =   1560
            Width           =   1920
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "f = 1.44/(R1 + R2 + R2) * C"
            Height          =   255
            Left            =   270
            TabIndex        =   42
            Top             =   795
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Duty Cycle     (low)"
            Height          =   390
            Left            =   105
            TabIndex        =   24
            Top             =   2970
            Width           =   780
         End
         Begin VB.Label lblDutyCycleL 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   975
            TabIndex        =   23
            Top             =   3045
            Width           =   1290
         End
         Begin VB.Label lblDutyCycleH 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   975
            TabIndex        =   22
            Top             =   2595
            Width           =   1290
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Duty Cycle      (high)"
            Height          =   450
            Left            =   105
            TabIndex        =   21
            Top             =   2535
            Width           =   885
         End
         Begin VB.Label lblOffTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            TabIndex        =   20
            Top             =   1935
            Width           =   1320
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Off Time:      (sec)"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   135
            TabIndex        =   19
            Top             =   1860
            Width           =   720
         End
         Begin VB.Label lblOnTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   960
            TabIndex        =   18
            Top             =   1185
            Width           =   1320
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "On Time:     (sec)"
            Height          =   405
            Left            =   150
            TabIndex        =   17
            Top             =   1155
            Width           =   690
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency (In Hertz)"
            Height          =   255
            Left            =   570
            TabIndex        =   16
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label lblFreq 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   570
            TabIndex        =   15
            Top             =   465
            Width           =   1470
         End
      End
      Begin VB.PictureBox picAstable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5580
         Left            =   105
         Picture         =   "Form1.frx":1AE9
         ScaleHeight     =   5550
         ScaleWidth      =   5955
         TabIndex        =   4
         Top             =   270
         Width           =   5985
         Begin VB.ComboBox cboC 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   705
            TabIndex        =   7
            Text            =   "1 ufd"
            Top             =   3690
            Width           =   1020
         End
         Begin VB.ComboBox cboR2 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   870
            TabIndex        =   6
            Text            =   "470K"
            Top             =   2355
            Width           =   840
         End
         Begin VB.ComboBox cboR1 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   855
            TabIndex        =   5
            Text            =   "47K"
            Top             =   990
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   6
            Visible         =   0   'False
            X1              =   5505
            X2              =   5445
            Y1              =   2970
            Y2              =   2820
         End
         Begin VB.Line Line1 
            Index           =   5
            Visible         =   0   'False
            X1              =   5775
            X2              =   5925
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Line Line1 
            Index           =   4
            Visible         =   0   'False
            X1              =   5370
            X2              =   5265
            Y1              =   3255
            Y2              =   3315
         End
         Begin VB.Line Line1 
            Index           =   3
            Visible         =   0   'False
            X1              =   5715
            X2              =   5835
            Y1              =   3360
            Y2              =   3480
         End
         Begin VB.Line Line1 
            Index           =   2
            Visible         =   0   'False
            X1              =   5700
            X2              =   5805
            Y1              =   3000
            Y2              =   2835
         End
         Begin VB.Line Line1 
            Index           =   1
            Visible         =   0   'False
            X1              =   5355
            X2              =   5190
            Y1              =   3105
            Y2              =   3000
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  'Solid
            Height          =   285
            Left            =   5445
            Shape           =   3  'Circle
            Top             =   3060
            Width           =   285
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Note: Normally R1 is 1/10 value of R2."
            Height          =   465
            Left            =   2205
            TabIndex        =   50
            Top             =   1215
            Width           =   1500
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "4.5 to 15 vdc"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   26
            Top             =   195
            Width           =   1395
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Output"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5205
            TabIndex        =   25
            Top             =   2280
            Width           =   690
         End
         Begin VB.Label Lb6 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   13
            Top             =   2655
            Width           =   315
         End
         Begin VB.Label Lb5 
            Appearance      =   0  'Flat
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   12
            Top             =   2505
            Width           =   315
         End
         Begin VB.Label Lb4 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   11
            Top             =   2355
            Width           =   315
         End
         Begin VB.Label Lb3 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   10
            Top             =   1410
            Width           =   315
         End
         Begin VB.Label Lb2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   9
            Top             =   1245
            Width           =   315
         End
         Begin VB.Label Lb1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   8
            Top             =   1095
            Width           =   315
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   6060
      Left            =   150
      TabIndex        =   2
      Top             =   1005
      Visible         =   0   'False
      Width           =   8700
      Begin VB.Frame Frame5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5610
         Left            =   6225
         TabIndex        =   29
         Top             =   255
         Width           =   2400
         Begin VB.PictureBox picMS_Output 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   945
            Left            =   120
            ScaleHeight     =   63
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   142
            TabIndex        =   49
            Top             =   3525
            Width           =   2130
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   30
               TabIndex        =   52
               Top             =   750
               Width           =   90
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Vcc"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Left            =   15
               TabIndex        =   51
               Top             =   -30
               Width           =   240
            End
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "high = 1.1 * R * C"
            Height          =   195
            Left            =   900
            TabIndex        =   45
            Top             =   750
            Width           =   1290
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "(seconds)"
            Height          =   195
            Left            =   1155
            TabIndex        =   41
            Top             =   195
            Width           =   690
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Output    (high)"
            Height          =   390
            Left            =   240
            TabIndex        =   40
            Top             =   375
            Width           =   600
         End
         Begin VB.Label lblMonoH 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   945
            TabIndex        =   39
            Top             =   420
            Width           =   1170
         End
      End
      Begin VB.PictureBox picMono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5580
         Left            =   105
         Picture         =   "Form1.frx":777DD
         ScaleHeight     =   5550
         ScaleWidth      =   5955
         TabIndex        =   28
         Top             =   270
         Width           =   5985
         Begin VB.ComboBox cboC2 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2580
            TabIndex        =   32
            Text            =   "1 ufd"
            Top             =   3915
            Width           =   1095
         End
         Begin VB.ComboBox cboR3 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2565
            TabIndex        =   31
            Text            =   "47K"
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   885
            TabIndex        =   38
            Top             =   1395
            Width           =   300
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   885
            TabIndex        =   37
            Top             =   1230
            Width           =   300
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   885
            TabIndex        =   36
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label Lb9 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   35
            Top             =   1395
            Width           =   315
         End
         Begin VB.Label Lb8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C000C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   34
            Top             =   1230
            Width           =   315
         End
         Begin VB.Label Lb7 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   1845
            TabIndex        =   33
            Top             =   1080
            Width           =   315
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   150
      TabIndex        =   53
      Top             =   1005
      Visible         =   0   'False
      Width           =   8685
      Begin VB.Image Image3 
         Height          =   6480
         Left            =   330
         Picture         =   "Form1.frx":ED4D1
         Stretch         =   -1  'True
         Top             =   -315
         Width           =   8100
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   6075
      Left            =   150
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   8700
      Begin VB.PictureBox picDS 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8F8F8&
         ForeColor       =   &H80000008&
         Height          =   5940
         Left            =   30
         ScaleHeight     =   5910
         ScaleWidth      =   8610
         TabIndex        =   30
         Top             =   105
         Width           =   8640
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0E0E0&
            Height          =   5865
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   30
            Visible         =   0   'False
            Width           =   6810
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Pin Description Show"
            Height          =   750
            Left            =   7140
            TabIndex        =   54
            Top             =   4800
            Width           =   1035
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   4185
            Left            =   135
            Picture         =   "Form1.frx":EFB60
            ScaleHeight     =   4185
            ScaleWidth      =   6375
            TabIndex        =   56
            Top             =   1680
            Width           =   6375
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   2115
            Left            =   6675
            Picture         =   "Form1.frx":F3401
            Stretch         =   -1  'True
            Top             =   150
            Width           =   2310
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   1620
            Left            =   1365
            Picture         =   "Form1.frx":F4EEA
            Top             =   60
            Width           =   4395
         End
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "555  TIMER CALCULATOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2400
      TabIndex        =   27
      Top             =   75
      Width           =   4185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ken Foster
'May 2010

Option Explicit
Dim co1 As Long
Dim co2 As Long
Dim co3 As Long
Dim DutyCycleH As Single
Dim DutyCycleL As Single

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Load()
    sTab1.List(0) = "Astable"
    sTab1.List(1) = "Monostable"
    sTab1.List(2) = "Diagrams"
    sTab1.List(3) = "Schematic"
    
    'load values into combo boxes
    LoadCombo 1, cboR1
    cboR1.ListIndex = 2
    cboR1.Text = "4.7K"
   
    LoadCombo 1, cboR2
    cboR2.ListIndex = 7
    cboR2.Text = "470K"
    
    LoadCombo 1, cboR3
    cboR3.ListIndex = 12
    cboR3.Text = "470K"
    
    LoadCombo 2, cboC
    cboC.ListIndex = 15
    cboC.Text = "1 ufd"
    
    LoadCombo 2, cboC2
    cboC2.ListIndex = 15
    cboC2.Text = "1 ufd"
    
    'pre calculate some values
    CalculateValues 1
    CalculateValues 2
    cboR1_Click
End Sub

Private Sub Command1_Click()
   Text1.Visible = Not Text1.Visible
   If Text1.Visible = False Then
      Command1.Caption = "Pin Description Show"
      Exit Sub
   Else
      Command1.Caption = "Pin Description Hide"
   End If
   Call GetTextFromFile(App.Path & "/555Text.txt", Text1)
End Sub
    
Private Sub Check1_Click()
   If Check1.value = Checked Then
      Timer1.Enabled = True
   Else
      Timer1.Enabled = False
   End If
End Sub

Private Sub CalculateValues(cv As Integer)
Dim R1 As Single
Dim R2 As Single
Dim C1 As Single
Dim MonoL As Single
Dim MonoH As Single
Dim rMono As Single
Dim OnTime As Single
Dim OffTime As Single
Dim rOnOff As Single

    Select Case cv
       Case 1   'Astable
          GetCompValue R1, cboR1
          GetCompValue R2, cboR2
          GetCompValue C1, cboC
    
          lblFreq.Caption = Format(1.44 / ((R1 + R2 + R2) * C1), "######0.###")
          lblOnTime.Caption = Format(0.693 * (R1 + R2) * C1, "##0.####")
         OnTime = Val(lblOnTime.Caption)
          lblOffTime.Caption = Format(0.693 * R2 * C1, "##0.####")
         OffTime = Val(lblOffTime.Caption)
    
         If OnTime > OffTime Then
            rOnOff = Format((OnTime / OffTime) / 2 * 100, "##.##")
            lblDutyCycleH.Caption = rOnOff
            lblDutyCycleL.Caption = 100 - rOnOff
         Else
            rOnOff = Format((OffTime / OnTime) / 2 * 100, "##.##")
            lblDutyCycleL.Caption = rOnOff
            lblDutyCycleH.Caption = 100 - rOnOff
         End If
         If OnTime + OffTime * 10 < 10 Then
           Timer1.Interval = 10
         Else
           Timer1.Interval = OnTime + OffTime * 10
         End If
         Plot 1
       Case 2    ' Monostable
          GetCompValue R1, cboR3
          GetCompValue C1, cboC2
    
          lblMonoH.Caption = Format(1.1 * R1 * C1, "###.######")
          MonoH = Val(lblMonoH.Caption)
          MonoL = 100 - MonoH
          picMS_Output.Cls
          If MonoH > MonoL Then
             lblMonoH.Caption = ""
             Exit Sub
          Else
             rMono = (MonoH / MonoL) / 2 * 100
             lblDutyCycleH.Caption = rMono
             lblDutyCycleL.Caption = 100 - rMono
             DutyCycleL = Format(MonoL, "##0.#####")
             DutyCycleH = Format(100 - DutyCycleL, "##0.#####")
          End If
          Plot 2
       End Select
End Sub

Private Sub Plot(ptg As Integer)
Dim wP As Integer
Dim wL As Integer
Dim dLen As Integer

   Select Case ptg
      Case 1     'Astable
         dLen = Val(lblDutyCycleH.Caption) + Val(lblDutyCycleL.Caption)
         pic1.Cls
        'draw graph
         For wP = 0 To (picAS_output.ScaleWidth * 2) Step dLen
            pic1.Line (wP, 2)-(wP + Val(lblDutyCycleH.Caption), 2), vbGreen
            pic1.Line (wP + Val(lblDutyCycleH.Caption), 2)-(wP + Val(lblDutyCycleH.Caption), 35), vbGreen
            pic1.Line (wP + Val(lblDutyCycleH.Caption), 35)-(wP + Val(lblDutyCycleH.Caption) + Val(lblDutyCycleL.Caption), 35), vbGreen
            pic1.Line (wP + Val(lblDutyCycleH.Caption) + Val(lblDutyCycleL.Caption), 35)-(wP + Val(lblDutyCycleH.Caption) + Val(lblDutyCycleL.Caption), 2), vbGreen
         Next wP
      Case 2     'Monostable
         dLen = DutyCycleH + DutyCycleL
         picMS_Output.Cls
         'draw graph
         For wP = -30 To picMS_Output.ScaleWidth Step dLen
            picMS_Output.Line (wP, 10)-(wP + DutyCycleH, 10), vbGreen
            picMS_Output.Line (wP + DutyCycleH, 10)-(wP + DutyCycleH, 50), vbGreen
            picMS_Output.Line (wP + DutyCycleH, 50)-(wP + DutyCycleH + DutyCycleL, 50), vbGreen
            picMS_Output.Line (wP + DutyCycleH + DutyCycleL, 50)-(wP + DutyCycleH + DutyCycleL, 10), vbGreen
        Next wP
   End Select
End Sub

Private Sub cboC_Click()
    CalculateValues 1
End Sub

Private Sub cboC2_Click()
   CalculateValues 2
End Sub

Private Sub cboR1_Click()
    DrawResistorColors cboR1.Text
    Lb1.BackColor = co1
    Lb2.BackColor = co2
    Lb3.BackColor = co3
    CalculateValues 1
End Sub

Private Sub cboR2_Click()
    DrawResistorColors cboR2.Text
    Lb4.BackColor = co1
    Lb5.BackColor = co2
    Lb6.BackColor = co3
    CalculateValues 1
End Sub

Private Sub cboR3_Click()
    DrawResistorColors cboR3.Text
    Lb7.BackColor = co1
    Lb8.BackColor = co2
    Lb9.BackColor = co3
    CalculateValues 2
End Sub

Private Sub sTab1_Click()
    If sTab1.ListIndex = 0 Then
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        Frame6.Visible = False
        sTab1.BackColorSelected = Frame1.BackColor
    End If
    If sTab1.ListIndex = 1 Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame6.Visible = False
        sTab1.BackColorSelected = Frame2.BackColor
    End If
    If sTab1.ListIndex = 2 Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
        Frame6.Visible = False
        sTab1.BackColorSelected = Frame3.BackColor
    End If
    If sTab1.ListIndex = 3 Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame6.Visible = True
        sTab1.BackColorSelected = Frame4.BackColor
    End If
End Sub

Private Function GetTextFromFile(txtFile, txtopen As TextBox)
    Dim sfile As String
    Dim nfile As Integer
    nfile = FreeFile
    sfile = txtFile
    Open sfile For Input As nfile
    txtopen = Input(LOF(nfile), nfile)
    Close nfile
End Function

Private Sub Timer1_Timer()
Dim LedOn As Long
Dim X As Integer
   pic1.Left = pic1.Left + 1
   If pic1.Left = 0 Then pic1.Left = -100
   LedOn = GetPixel(pic1.hdc, -1 - pic1.Left, 2)
   If LedOn = vbGreen Then
      Shape1.FillColor = vbRed
      For X = 1 To 6
         Line1(X).Visible = True
      Next X
   Else
      Shape1.FillColor = vbBlack
      For X = 1 To 6
         Line1(X).Visible = False
      Next X
   End If
End Sub

Private Sub LoadCombo(Table As Integer, cbo As ComboBox)
Select Case Table
   Case 1
      cbo.AddItem "1K"
      cbo.AddItem "2.2K"
      cbo.AddItem "4.7K"
      cbo.AddItem "6.8K"
      cbo.AddItem "8.2K"
      cbo.AddItem "10K"
      cbo.AddItem "22K"
      cbo.AddItem "47K"
      cbo.AddItem "68K"
      cbo.AddItem "82K"
      cbo.AddItem "100K"
      cbo.AddItem "220K"
      cbo.AddItem "470K"
      cbo.AddItem "680K"
      cbo.AddItem "820K"
      cbo.AddItem "1 M"
      cbo.AddItem "2.2 M"
      cbo.AddItem "4.7 M"
      cbo.AddItem "6.8 M"
      cbo.AddItem "8.2 M"
      cbo.AddItem "10 M"
   
   Case 2
      cbo.AddItem "1 pf"
      cbo.AddItem "2.2 pf"
      cbo.AddItem "4.7 pf"
      cbo.AddItem "10 pf"
      cbo.AddItem "22 pf"
      cbo.AddItem "47 pf"
      cbo.AddItem "100 pf"
      cbo.AddItem "220 pf"
      cbo.AddItem "470 pf"
      cbo.AddItem ".001 ufd"
      cbo.AddItem ".0022 ufd"
      cbo.AddItem ".0047 ufd"
      cbo.AddItem ".01 ufd"
      cbo.AddItem ".022 ufd"
      cbo.AddItem ".047 ufd"
      cbo.AddItem ".1 ufd"
      cbo.AddItem ".22 ufd"
      cbo.AddItem ".47 ufd"
      cbo.AddItem "1 ufd"
      cbo.AddItem "2.2 ufd"
      cbo.AddItem "4.7 ufd"
      cbo.AddItem "10 ufd"
      cbo.AddItem "22 ufd"
      cbo.AddItem "47 ufd"
      cbo.AddItem "100 ufd"
      cbo.AddItem "220 ufd"
      cbo.AddItem "470 ufd"
      cbo.AddItem "1000 ufd"
   End Select
End Sub

Private Function GetCompValue(value As Single, vcbo As ComboBox)

   Select Case vcbo.Text
        Case "1K": value = 1000
        Case "2.2K": value = 2200
        Case "4.7K": value = 4700
        Case "6.8K": value = 6800
        Case "8.2K": value = 8200
        Case "10K": value = 10000
        Case "22K": value = 22000
        Case "47K": value = 47000
        Case "68K": value = 68000
        Case "82K": value = 82000
        Case "100K": value = 100000
        Case "220K": value = 220000
        Case "470K": value = 470000
        Case "680K": value = 680000
        Case "820K": value = 820000
        Case "1 M": value = 1000000
        Case "2.2 M": value = 2200000
        Case "4.7 M": value = 4700000
        Case "6.8 M": value = 6800000
        Case "8.2 M": value = 8200000
        Case "10 M": value = 10000000
        
        Case "1 pf": value = 0.000000000001
        Case "2.2 pf": value = 0.0000000000022
        Case "4.7 pf": value = 0.0000000000047
        Case "10 pf": value = 0.00000000001
        Case "22 pf": value = 0.000000000022
        Case "47 pf": value = 0.000000000047
        Case "100 pf": value = 0.0000000001
        Case "220 pf": value = 0.00000000022
        Case "470 pf": value = 0.00000000047
        Case ".001 ufd": value = 0.000000001
        Case ".0022 ufd": value = 0.0000000022
        Case ".0047 ufd": value = 0.0000000047
        Case ".01 ufd": value = 0.00000001
        Case ".022 ufd": value = 0.000000022
        Case ".047 ufd": value = 0.000000047
        Case ".1 ufd":   value = 0.0000001
        Case ".22 ufd": value = 0.00000022
        Case ".47 ufd": value = 0.00000047
        Case "1 ufd": value = 0.000001
        Case "2.2 ufd": value = 0.0000022
        Case "4.7 ufd": value = 0.0000047
        Case "10 ufd": value = 0.00001
        Case "22 ufd": value = 0.000022
        Case "47 ufd": value = 0.000047
        Case "100 ufd": value = 0.0001
        Case "220 ufd": value = 0.00022
        Case "470 ufd": value = 0.00047
        Case "1000 ufd": value = 0.001
    End Select
End Function

Private Sub DrawResistorColors(value As String)
    Select Case value
        Case "1K"
            co1 = &H4080&
            co2 = vbBlack
            co3 = vbRed
        Case "2.2K"
            co1 = vbRed
            co2 = vbRed
            co3 = vbRed
        Case "4.7K"
            co1 = vbYellow
            co2 = &HC000C0
            co3 = vbRed
        Case "6.8K"
            co1 = vbBlue
            co2 = &H80000010
            co3 = vbRed
        Case "8.2K"
            co1 = &H80000010
            co2 = vbRed
            co3 = vbRed
        Case "10K"
            co1 = &H4080&
            co2 = vbBlack
            co3 = &H80FF&
        Case "22K"
            co1 = vbRed
            co2 = vbRed
            co3 = &H80FF&
        Case "47K"
            co1 = vbYellow
            co2 = &HC000C0
            co3 = &H80FF&
        Case "68K"
            co1 = vbBlue
            co2 = &H80000010
            co3 = &H80FF&
         Case "82K"
            co1 = &H80000010
            co2 = vbRed
            co3 = &H80FF&
        Case "100K"
            co1 = &H4080&
            co2 = vbBlack
            co3 = vbYellow
        Case "220K"
            co1 = vbRed
            co2 = vbRed
            co3 = vbYellow
        Case "470K"
            co1 = vbYellow
            co2 = &HC000C0
            co3 = vbYellow
        Case "680K"
            co1 = vbBlue
            co2 = &H80000010
            co3 = vbYellow
        Case "820K"
            co1 = &H80000010
            co2 = vbRed
            co3 = vbYellow
        Case "1 M"
            co1 = &H4080&
            co2 = vbBlack
            co3 = vbGreen
        Case "2.2 M"
            co1 = vbRed
            co2 = vbRed
            co3 = vbGreen
        Case "4.7 M"
            co1 = vbYellow
            co2 = &HC000C0
            co3 = vbGreen
        Case "6.8 M"
            co1 = vbBlue
            co2 = &H80000010
            co3 = vbGreen
        Case "8.2 M"
            co1 = &H80000010
            co2 = vbRed
            co3 = vbGreen
        Case "10 M"
            co1 = &H4080&
            co2 = vbBlack
            co3 = vbBlue
    End Select
End Sub
