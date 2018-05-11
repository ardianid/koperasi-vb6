VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form Karyawan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATA KARYAWAN           "
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Counter 
      Height          =   3975
      Left            =   -2160
      TabIndex        =   119
      Top             =   9120
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   7011
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":08CA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":08E6
      Childs          =   "Karyawan.frx":0992
      Begin VB.TextBox Txt_Cr_Counter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   126
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Txt_Cr_Counter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   125
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   240
         TabIndex        =   120
         Top             =   480
         Width           =   6495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Counter 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":09AE
         TabIndex        =   121
         Top             =   1080
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   480
         TabIndex        =   124
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   2880
         TabIndex        =   123
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   122
         Top             =   240
         Width           =   1065
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Bagian 
      Height          =   2295
      Left            =   2400
      TabIndex        =   79
      Top             =   9120
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":3443
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":345F
      Childs          =   "Karyawan.frx":350B
      Begin VB.TextBox Txt_Cr_Bagian 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   84
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_Cr_Bagian 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   83
         Top             =   240
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Bagian 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":3527
         TabIndex        =   80
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   61
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         Height          =   195
         Index           =   60
         Left            =   2400
         TabIndex        =   81
         Top             =   240
         Width           =   480
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3855
      Left            =   -5280
      TabIndex        =   59
      Top             =   2160
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":5E7F
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":5E9B
      Childs          =   "Karyawan.frx":5F47
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   106
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   64
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   63
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":5F63
         TabIndex        =   60
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   105
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   41
         Left            =   2640
         TabIndex        =   62
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   40
         Left            =   600
         TabIndex        =   61
         Top             =   600
         Width           =   360
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Hapus 
      Height          =   3615
      Left            =   -5280
      TabIndex        =   53
      Top             =   1680
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":8ED3
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":8EEF
      Childs          =   "Karyawan.frx":8F9B
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   104
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   58
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   57
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Hapus 
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":8FB7
         TabIndex        =   54
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   103
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   39
         Left            =   480
         TabIndex        =   56
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   38
         Left            =   2640
         TabIndex        =   55
         Top             =   600
         Width           =   405
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   3735
      Left            =   -5280
      TabIndex        =   47
      Top             =   2520
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6588
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":BF26
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":BF42
      Childs          =   "Karyawan.frx":BFEE
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   102
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   49
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   48
         Top             =   600
         Width           =   1215
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Karyawan.frx":C00A
         TabIndex        =   50
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   101
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   37
         Left            =   2760
         TabIndex        =   52
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   36
         Left            =   480
         TabIndex        =   51
         Top             =   600
         Width           =   360
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Jabatan 
      Height          =   2295
      Left            =   -2520
      TabIndex        =   40
      Top             =   8760
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":EF79
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":EF95
      Childs          =   "Karyawan.frx":F041
      Begin VB.TextBox Txt_Cr_Jabatan 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   42
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_Cr_Jabatan 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Jabatan 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":F05D
         TabIndex        =   43
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   195
         Index           =   35
         Left            =   2400
         TabIndex        =   45
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   360
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Pendidikan 
      Height          =   2295
      Left            =   -3600
      TabIndex        =   34
      Top             =   8520
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":119B6
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":119D2
      Childs          =   "Karyawan.frx":11A7E
      Begin VB.TextBox Txt_Cr_Pendidikan 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Pendidikan 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   2175
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Pendidikan 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":11A9A
         TabIndex        =   37
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan"
         Height          =   195
         Index           =   32
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   765
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Status_P 
      Height          =   2295
      Left            =   3840
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":143FA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":14416
      Childs          =   "Karyawan.frx":144C2
      Begin VB.TextBox Txt_Cr_Status 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Txt_Cr_Status 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Status 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":144DE
         TabIndex        =   31
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   30
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   465
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Agama 
      Height          =   2295
      Left            =   -3480
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   4048
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Karyawan.frx":16E42
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Karyawan.frx":16E5E
      Childs          =   "Karyawan.frx":16F0A
      Begin VB.TextBox Txt_Cr_Agama 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Txt_Cr_Agama 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Agama 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "Karyawan.frx":16F26
         TabIndex        =   25
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
         Height          =   195
         Index           =   29
         Left            =   2400
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   113
         Top             =   10920
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Cmd_Browse_Counter 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Left            =   4200
            TabIndex        =   117
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Lbl_Kode_Counter 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   118
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Lbl_Nama_Counter 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   116
            Top             =   360
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   22
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   1200
            TabIndex        =   114
            Top             =   360
            Width           =   60
         End
      End
      Begin VB.TextBox Txt_Kode_Agama 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   100
         Top             =   2280
         Width           =   495
      End
      Begin VB.ComboBox Cbo_Agama 
         Height          =   315
         ItemData        =   "Karyawan.frx":1987D
         Left            =   2040
         List            =   "Karyawan.frx":1987F
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Frame Frame_Nav 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   93
         Top             =   6000
         Width           =   2175
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "SansSerif"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   1560
            TabIndex        =   97
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "SansSerif"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   1080
            TabIndex        =   96
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "SansSerif"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   95
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Cmd_Navigasi 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "SansSerif"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   85
         Top             =   6000
         Width           =   4455
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   3480
            TabIndex        =   90
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Daftar"
            Height          =   495
            Left            =   2640
            TabIndex        =   89
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Hapus 
            Caption         =   "&Hapus"
            Height          =   495
            Left            =   1800
            TabIndex        =   88
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   495
            Left            =   960
            TabIndex        =   87
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   495
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   495
            Left            =   960
            TabIndex        =   91
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   1080
         TabIndex        =   69
         Top             =   7920
         Visible         =   0   'False
         Width           =   4695
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   107
            Top             =   600
            Width           =   4455
            Begin VB.OptionButton Opt_Hari 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Per&Hari"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   112
               Top             =   650
               Width           =   975
            End
            Begin VB.OptionButton Opt_Bulan 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Per&Bulan"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   111
               Top             =   650
               Width           =   975
            End
            Begin TDBNumber6Ctl.TDBNumber TDB_Gaji 
               Height          =   320
               Left            =   1320
               TabIndex        =   108
               Top             =   240
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   564
               Calculator      =   "Karyawan.frx":19881
               Caption         =   "Karyawan.frx":198A1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Karyawan.frx":1990D
               Keys            =   "Karyawan.frx":1992B
               Spin            =   "Karyawan.frx":19975
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999
               MinValue        =   -999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   1
               Value           =   0
               MaxValueVT      =   1028849669
               MinValueVT      =   1598423045
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   20
               Left            =   1200
               TabIndex        =   110
               Top             =   240
               Width           =   60
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gaji Pokok"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   26
               Left            =   120
               TabIndex        =   109
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.TextBox Txt_Ket 
            Height          =   765
            Left            =   1440
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   2400
            Width           =   3015
         End
         Begin TDBNumber6Ctl.TDBNumber TDB_Tunjangan 
            Height          =   320
            Left            =   1440
            TabIndex        =   72
            Top             =   1680
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   564
            Calculator      =   "Karyawan.frx":1999D
            Caption         =   "Karyawan.frx":199BD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Karyawan.frx":19A29
            Keys            =   "Karyawan.frx":19A47
            Spin            =   "Karyawan.frx":19A91
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###;;0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2089877505
            Value           =   0
            MaxValueVT      =   1028849669
            MinValueVT      =   1598423045
         End
         Begin TDBNumber6Ctl.TDBNumber TDB_Uang_Makan 
            Height          =   320
            Left            =   1440
            TabIndex        =   75
            Top             =   2040
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   564
            Calculator      =   "Karyawan.frx":19AB9
            Caption         =   "Karyawan.frx":19AD9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Karyawan.frx":19B45
            Keys            =   "Karyawan.frx":19B63
            Spin            =   "Karyawan.frx":19BAD
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###;;0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2089877505
            Value           =   0
            MaxValueVT      =   1028849669
            MinValueVT      =   1598423045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   59
            Left            =   1320
            TabIndex        =   77
            Top             =   2400
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ket"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   58
            Left            =   240
            TabIndex        =   76
            Top             =   2400
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   49
            Left            =   1320
            TabIndex        =   74
            Top             =   2040
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uang Makan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   48
            Left            =   240
            TabIndex        =   73
            Top             =   2040
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   47
            Left            =   1320
            TabIndex        =   71
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tunjangan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   46
            Left            =   240
            TabIndex        =   70
            Top             =   1680
            Width           =   870
         End
      End
      Begin VB.TextBox Txt_Kode_Jenis_Kelamin 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   67
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox Cbo_Jenis_Kelamin 
         Height          =   315
         ItemData        =   "Karyawan.frx":19BD5
         Left            =   2040
         List            =   "Karyawan.frx":19BD7
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1920
         Width           =   2415
      End
      Begin TDBDate6Ctl.TDBDate TDB_Tgl_Lhr 
         Height          =   315
         Left            =   5640
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "Karyawan.frx":19BD9
         Caption         =   "Karyawan.frx":19CF1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Karyawan.frx":19D5D
         Keys            =   "Karyawan.frx":19D7B
         Spin            =   "Karyawan.frx":19DD9
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   1863103
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "05/01/2007"
         ValidateMode    =   0
         ValueVT         =   2010382343
         Value           =   39087
         CenturyMode     =   0
      End
      Begin VB.TextBox Txt_Tempat_Lhr 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Txt_Kodepos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Txt_Alamat_3 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   3360
         Width           =   5655
      End
      Begin VB.TextBox Txt_Alamat_2 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox Txt_Alamat_1 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox Txt_Nama 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox Txt_Kode 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate TDB_Tgl_Masuk 
         Height          =   315
         Left            =   1560
         TabIndex        =   127
         Top             =   4800
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   547
         Calendar        =   "Karyawan.frx":19E01
         Caption         =   "Karyawan.frx":19F19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Karyawan.frx":19F85
         Keys            =   "Karyawan.frx":19FA3
         Spin            =   "Karyawan.frx":1A001
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   1863103
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "05/01/2007"
         ValidateMode    =   0
         ValueVT         =   2010382343
         Value           =   39087
         CenturyMode     =   0
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   5640
         TabIndex        =   20
         Top             =   0
         Width           =   2535
         Begin VB.Label Lbl_Umur 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Umur Karyawan :"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "No. Telp"
         Height          =   855
         Left            =   360
         TabIndex        =   11
         Top             =   3960
         Width           =   7575
         Begin VB.TextBox Txt_Telp_Hp 
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Txt_Telp_Rumah 
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp Hp :"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   13
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp Rumah :"
            Height          =   195
            Index           =   8
            Left            =   15
            TabIndex        =   12
            Top             =   120
            Width           =   1185
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Masuk :"
         Height          =   195
         Index           =   50
         Left            =   720
         TabIndex        =   128
         Top             =   4800
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agama :"
         Height          =   195
         Index           =   18
         Left            =   960
         TabIndex        =   98
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7320
         TabIndex        =   68
         Top             =   5760
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin :"
         Height          =   195
         Index           =   42
         Left            =   495
         TabIndex        =   65
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lhr :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   4920
         TabIndex        =   18
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lhr :"
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos :"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   9
         Top             =   3720
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   4
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode :"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   465
      End
   End
End
Attribute VB_Name = "Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
Dim Arr_Rubah As New XArrayDB
Dim Arr_Hapus As New XArrayDB
Dim arr_daftar As New XArrayDB

Private Sub Kosong_Rubah()

    Arr_Rubah.ReDim 0, 0, 0, 0
    Arr_Rubah.ReDim 1, 1, 1, 1
    Grid_Rubah.ReBind
    Grid_Rubah.Refresh
    
End Sub

Private Sub Kosong_Hapus()
    Arr_Hapus.ReDim 0, 0, 0, 0
    Arr_Hapus.ReDim 1, 1, 1, 1
    Grid_Hapus.ReBind
    Grid_Hapus.Refresh
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    arr_daftar.ReDim 1, 1, 1, 1
    Grid_Daftar.ReBind
    Grid_Daftar.Refresh
End Sub


Private Sub Cbo_Agama_Change()
    With Cbo_Agama
        Txt_Kode_Agama.Text = Left(.Text, 2)
    End With
End Sub

Private Sub Cbo_Agama_Click()
    Cbo_Agama_Change
End Sub

Private Sub Cbo_Agama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Alamat_1.SetFocus
End Sub

Private Sub Cbo_Jenis_Kelamin_Change()
    With Cbo_Jenis_Kelamin
        Txt_Kode_Jenis_Kelamin.Text = Left(.Text, 2)
    End With
End Sub

Private Sub Cbo_Jenis_Kelamin_Click()
    Cbo_Jenis_Kelamin_Change
End Sub

Private Sub cbo_jenis_kelamin_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Cbo_Agama.SetFocus
End If

End Sub

Private Sub Cmd_Batal_Click()
    
    If rubah <> True Then Lbl_Umur.Caption = "0 Thn"
    
    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        
        Cmd_Tambah.Enabled = True
    
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
        Cmd_Hapus.Enabled = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        
        If TypeOf n Is TDBNumber Then
            n.Enabled = False
        End If
        
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
        If TypeOf n Is OptionButton Then n.Enabled = False

Next

Set n = Nothing

 If Cmd_Tambah.Enabled = True Then Cmd_Tambah.SetFocus

 Txt_Cr_Daftar_KeyUp 0, 0, 0
 Cmd_Navigasi_Click 3
    
End Sub

Private Sub Cmd_Browse_Counter_Click()

With TDB_Counter
    .Left = 2640
    .Top = 1200
    
    If .Visible = False Then
    
    Txt_Cr_Counter(0).Text = ""
    Txt_Cr_Counter(1).Text = ""
    
    Txt_Cr_Counter_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Counter(0).SetFocus
    
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub Cmd_Daftar_Click()

Frame_Nav.Enabled = False
With TDB_Daftar

If .Visible = False Then
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Daftar(0).Text = ""
    Txt_Cr_Daftar(1).Text = ""
    
    Txt_Cr_Daftar_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Daftar(0).SetFocus
    
Else
    .Visible = False
End If

End With

End Sub

Private Sub Cmd_Hapus_Click()

Frame_Nav.Enabled = False
With TDB_Hapus

If .Visible = False Then
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Hapus(0).Text = ""
    Txt_Cr_Hapus(1).Text = ""
    
    Txt_Cr_Hapus_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Hapus(0).SetFocus
    
Else
    .Visible = False
End If

End With

End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Navigasi_Click(Index As Integer)
On Error Resume Next

With Rs_Nav
Select Case Index
    Case 0
        .MoveFirst
    Case 1
        
        If .BOF Then .MoveFirst
        
        .MovePrevious
        
        If .BOF Then .MoveFirst
        
    Case 2
        
        If .EOF Then .MoveLast
        
        .MoveNext
        
        If .EOF Then .MoveLast
        
    Case 3
        
        .MoveLast
        
End Select
End With

isi_semua Rs_Nav

End Sub

Private Sub Cmd_Rubah_Click()

Frame_Nav.Enabled = False
With TDB_Rubah

If .Visible = False Then
    
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    Txt_Cr_Rubah(0).Text = ""
    Txt_Cr_Rubah(1).Text = ""
    
    Txt_Cr_Rubah_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Rubah(0).SetFocus
    
Else
    .Visible = False
End If

End With

End Sub

Private Sub Cmd_Simpan_Click()
On Error GoTo err_handler

Dim konfirm As Integer
            
'            Dim Gapok As Double
'            If TDB_Gaji.ValueIsNull Then
'                Gapok = 0
'            Else
'                Gapok = Replace(Trim(TDB_Gaji.Value), ",", "")
'            End If
'
'
'            Dim Tunj As Double
'            If TDB_Tunjangan.ValueIsNull Then
'               Tunj = 0
'            Else
'                Tunj = Replace(Trim(TDB_Tunjangan.Value), ",", "")
'            End If
'
'            Dim Uang_Mkan As Double
'            If TDB_Uang_Makan.ValueIsNull Then
'                Uang_Mkan = 0
'            Else
'                Uang_Mkan = Replace(Trim(TDB_Uang_Makan.Value), ",", "")
'            End If
            
kon.BeginTrans
Dim sql, sql1 As String
Dim rs As Recordset
Dim rs1 As Recordset
            
'Dim fl_gaji As String
'    If Opt_Bulan.Value = True Then
'        fl_gaji = "b"
'    ElseIf Opt_Hari.Value = True Then
'        fl_gaji = "h"
'    End If
            
If rubah = False Then
    
    If Txt_Kode.Text = "" Then
        konfirm = CInt(MsgBox("Kode karyawan tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        Txt_Kode.SetFocus
        
        On Error GoTo 0
        Exit Sub
    Else
    
    sql1 = "select Kode_Karyawan from Tb_Karyawan where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
    
    Set rs1 = New ADODB.Recordset
        rs1.Open sql1, kon
    
    With rs1
        If Not .EOF Then
            konfirm = CInt(MsgBox("Kode karyawan yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Kode.SetFocus
            
            kon.RollbackTrans
            On Error GoTo 0
            Exit Sub
        End If
    End With
        
        sql = "insert into Tb_Karyawan (Kode_Karyawan,Nama_Karyawan,Jenis_Kelamin,Agama,Alamat_1,Alamat_2,Alamat_3,Kode_Pos,No_Telp,No_Telp_HP,Tempat_Lhr,Tgl_Lhr,Tgl_Masuk,Jml_Hutang)"
        sql = sql & " values('" & Trim(Txt_Kode.Text) & "','" & Trim(Txt_Nama.Text) & "','" & Trim(Txt_Kode_Jenis_Kelamin.Text) & "','" & Trim(Txt_Kode_Agama.Text) & "','" & Trim(Txt_Alamat_1.Text) & "','" & Trim(Txt_Alamat_2.Text) & "','" & Trim(Txt_Alamat_3.Text) & "','" & Trim(Txt_Kodepos.Text) & "'"
        sql = sql & ",'" & Trim(Txt_Telp_Rumah.Text) & "','" & Trim(Txt_Telp_Hp.Text) & "','" & Trim(Txt_Tempat_Lhr.Text) & "','" & Format(Trim(TDB_Tgl_Lhr.Text), "yyyy/mm/dd") & "','" & Format(Trim(TDB_Tgl_Masuk.Text), "yyyy/mm/dd") & "',0)"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data karyawan telah disimpan ...", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Batal_Click
        
    End If
    
Else

    sql = "update Tb_Karyawan set Nama_Karyawan='" & Trim(Txt_Nama.Text) & "',Jenis_Kelamin='" & Trim(Txt_Kode_Jenis_Kelamin.Text) & "',Alamat_1='" & Trim(Txt_Alamat_1.Text) & "',Alamat_2='" & Trim(Txt_Alamat_2.Text) & "',Alamat_3='" & Trim(Txt_Alamat_3.Text) & "',"
    sql = sql & "Kode_Pos='" & Trim(Txt_Kodepos.Text) & "',No_Telp_Hp='" & Trim(Txt_Telp_Hp.Text) & "',No_Telp='" & Trim(Txt_Telp_Rumah.Text) & "',Agama='" & Trim(Txt_Kode_Agama.Text) & "',Tempat_Lhr='" & Trim(Txt_Tempat_Lhr.Text) & "',Tgl_Lhr='" & Format(Trim(TDB_Tgl_Lhr.Text), "yyyy/mm/dd") & "',"
    sql = sql & "Tgl_Masuk='" & Format(Trim(TDB_Tgl_Masuk.Text), "yyyy/mm/dd") & "' where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
        
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data karyawan telah dirubah ...", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Batal_Click
    
End If

rubah = False
On Error GoTo 0
Exit Sub

err_handler:
    
    kon.RollbackTrans
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear

End Sub

Private Sub Cmd_Tambah_Click()
    
    rubah = False
    
    Frame_Nav.Enabled = False
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
     Cmd_Rubah.Visible = False
     Cmd_Batal.Visible = True
     Cmd_Hapus.Enabled = False
     Cmd_Daftar.Enabled = False
     Cmd_Keluar.Enabled = False
        
     Txt_Kode.Text = ""
     Txt_Kode.Enabled = True
     Txt_Kode.SetFocus
        
        
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Cmd_Tambah.SetFocus
End Sub

Private Sub Form_Load()

Dim status As String
status = Buka_Koneksi
If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
        
        End
        Exit Sub
End If

rubah = False

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Tambah.Enabled = c_tambah
'    Cmd_Rubah.Enabled = c_rubah
'    Cmd_Hapus.Enabled = c_hapus

'' stop here ''


With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 350
End With

Dim n As Object
    For Each n In Me
    
        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
'        If TypeOf n Is CheckBox Then n.Enabled = False
'        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is TDBNumber Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

Set n = Nothing

Grid_Rubah.Array = Arr_Rubah
Grid_Hapus.Array = Arr_Hapus
Grid_Daftar.Array = arr_daftar


With Cbo_Agama
    .Clear
    .AddItem "01. Islam"
    .AddItem "02. Kristen"
    .AddItem "03. Hindu"
    .AddItem "04. Budha"
    .AddItem "05. Konghucu"
End With

atur_grid_transaksi

Isi_Combo

Cmd_Simpan.TabIndex = Txt_Ket.TabIndex + 1

Lbl_Umur.Caption = "0 Thn"

Txt_Cr_Daftar_KeyUp 0, 0, 0

Cmd_Navigasi_Click 3

End Sub

Sub Isi_Combo()
    With Cbo_Jenis_Kelamin
        .Clear
        .AddItem "01. Pria"
        .AddItem "02. Wanita"
    End With
End Sub

Sub atur_grid_transaksi()
    
    With TDB_Rubah
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With
    
    With TDB_Hapus
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With

    With TDB_Daftar
        .Left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
    End With

End Sub

Sub Isi_grid_transaksi(ByVal rec As Recordset, ByVal gridnya As Integer)
    
    Dim a As Long
    Dim kode, NAMA, alamat As String
        
        Select Case gridnya
            Case 0
                Kosong_Rubah
            Case 1
                Kosong_Hapus
            Case 2
                kosong_daftar
        End Select
        
        a = 1
        
        With rec
            
           Do While Not .EOF
            
           Select Case gridnya
            Case 0
                Arr_Rubah.ReDim 1, a, 0, Grid_Rubah.Columns.Count
                Grid_Rubah.ReBind
                Grid_Rubah.Refresh
             Case 1
                Arr_Hapus.ReDim 1, a, 0, Grid_Hapus.Columns.Count
                Grid_Hapus.ReBind
                Grid_Hapus.Refresh
             Case 2
                arr_daftar.ReDim 1, a, 0, Grid_Daftar.Columns.Count
                Grid_Daftar.ReBind
                Grid_Daftar.Refresh
           End Select
            
            kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
            NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
            alamat = IIf(Not IsNull(!alamat_1), !alamat_1, "")
            
            Select Case gridnya
                Case 0
                    Arr_Rubah(a, 0) = kode
                    Arr_Rubah(a, 1) = NAMA
                    Arr_Rubah(a, 2) = alamat
                Case 1
                    Arr_Hapus(a, 0) = kode
                    Arr_Hapus(a, 1) = NAMA
                    Arr_Hapus(a, 2) = alamat
                Case 2
                    arr_daftar(a, 0) = kode
                    arr_daftar(a, 1) = NAMA
                    arr_daftar(a, 2) = alamat
            End Select
            
           a = a + 1
           .MoveNext
           Loop
            
           Select Case gridnya
            Case 0
            
                Grid_Rubah.ReBind
                Grid_Rubah.Refresh
                
                Grid_Rubah.MoveFirst
                
            Case 1
                
                Grid_Hapus.ReBind
                Grid_Hapus.Refresh
                
                Grid_Hapus.MoveLast
                
            Case 2
                
                Grid_Daftar.ReBind
                Grid_Daftar.Refresh
                
                Grid_Daftar.MoveLast
                
           End Select
            
        End With
End Sub

Sub Atur_Tdb(ByVal TDB As TDBContainer3D, ByVal frme As Frame, ByVal comd As CommandButton)
    With TDB
        .Left = frme.Left + comd.Left - .Width / 2
        .Top = frme.Top + comd.Top + comd.Height
    End With
End Sub

Sub Atur_Tdb_Atas(ByVal TDB As TDBContainer3D, ByVal frme As Frame, ByVal comd As CommandButton)
    With TDB
        .Left = frme.Left + comd.Left - .Width / 2
        .Top = frme.Top + comd.Top - .Height
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Cmd_Keluar.Enabled = False Then
        Cancel = True
    Else
        Cancel = False

    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
        
    End If
End Sub

Private Sub Grid_Counter_DblClick()

If Grid_Counter.Row < 0 Then Exit Sub

    Lbl_Kode_Counter.Caption = Grid_Counter.Columns(0).Text
    Lbl_Nama_Counter.Caption = Grid_Counter.Columns(1).Text
    
    TDB_Counter.Visible = False
    Cmd_Simpan.SetFocus

End Sub

Private Sub Grid_Counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Counter_DblClick
    If KeyCode = vbKeyEscape Then TDB_Counter.Visible = False: Cmd_Browse_Counter.SetFocus
End Sub

Private Sub grid_daftar_DblClick()
    
    If arr_daftar.UpperBound(1) = 1 And arr_daftar(1, 1) = Empty Then Exit Sub

    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "Kode_Karyawan='" & arr_daftar(Grid_Daftar.Bookmark, 0) & "'"

    isi_semua Rs_Nav
    
    TDB_Daftar.Visible = False
    Frame_Nav.Enabled = True
    Cmd_Navigasi(0).SetFocus
    
End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Grid_Hapus_DblClick()
    
On Error GoTo err_handler
    
    If Arr_Hapus.UpperBound(1) = 1 And Arr_Hapus(1, 1) = Empty Then Exit Sub
    
    kon.BeginTrans
    
    If MsgBox("Yakin akan hapus : " & Arr_Hapus(Grid_Hapus.Bookmark, 0) & " ...?", vbYesNo + vbQuestion, "Hapus") = vbNo Then
        kon.RollbackTrans
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "delete from Tb_Karyawan where Kode_Karyawan='" & Arr_Hapus(Grid_Hapus.Bookmark, 0) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
        
        kon.CommitTrans
        Dim konfirm As Integer
            
            konfirm = CInt(MsgBox(Arr_Hapus(Grid_Hapus.Bookmark, 0) & " Berhasil dihapus", vbOKOnly + vbInformation, "Hapus"))
            
            Cmd_Batal_Click
        
        On Error GoTo 0
        Exit Sub
        
err_handler:
    
    kon.RollbackTrans
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
    
End Sub

Private Sub Grid_Hapus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Grid_Rubah_DblClick()

If Arr_Rubah.UpperBound(1) = 1 And Arr_Rubah(1, 1) = Empty Then Exit Sub

    Txt_Cr_Rubah(0).Text = Arr_Rubah(Grid_Rubah.Bookmark, 0)

    Txt_Cr_Rubah_KeyUp 0, 0, 0

    isi_semua Rs_Nav
    
    TDB_Rubah.Visible = False
        
        
    Dim n As Object
        For Each n In Me
                        If TypeOf n Is TextBox Then
                        
                         If Not (Left(UCase(n.Name), 9) = UCase("Txt_Kode_") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan" Or n.Name = "Txt_Kode") Then
                            n.Enabled = True
                         End If
                         
                        End If
            
            If TypeOf n Is TDBDate Then n.Enabled = True
            If TypeOf n Is TDBNumber Then n.Enabled = True
            If TypeOf n Is ComboBox Then n.Enabled = True
            If TypeOf n Is OptionButton Then n.Enabled = True
'            If TypeOf n Is OptionButton Then n.Enabled = True
'            If TypeOf n Is CheckBox Then n.Enabled = True
            
            If TypeOf n Is CommandButton Then
                If n.Caption = "..." Then
                    n.Enabled = True
                End If
            End If
            
        Next

    Cmd_Simpan.Enabled = True
    rubah = True
    
    Txt_Nama.SetFocus
    
End Sub

Sub isi_semua(ByVal rec As Recordset)
On Error Resume Next

    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        Txt_Kode = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
        Txt_Nama = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
        Txt_Tempat_Lhr = IIf(Not IsNull(!tempat_lhr), !tempat_lhr, "")
        TDB_Tgl_Lhr.Value = IIf(Not IsNull(!tgl_lhr), !tgl_lhr, Date)
        
        TDB_Tgl_Lhr_LostFocus
        
        Txt_Kode_Jenis_Kelamin = IIf(Not IsNull(!jenis_kelamin), !jenis_kelamin, "")
        
        Txt_Alamat_1 = IIf(Not IsNull(!alamat_1), !alamat_1, "")
        Txt_Alamat_2 = IIf(Not IsNull(!alamat_2), !alamat_2, "")
        Txt_Alamat_3 = IIf(Not IsNull(!alamat_3), !alamat_3, "")
        Txt_Kodepos = IIf(Not IsNull(!Kode_Pos), !Kode_Pos, "")
        Txt_Telp_Rumah = IIf(Not IsNull(!No_Telp), !No_Telp, "")
        Txt_Telp_Hp = IIf(Not IsNull(!No_Telp_Hp), !No_Telp_Hp, "")
        Txt_Kode_Agama = IIf(Not IsNull(!Agama), !Agama, "")
        TDB_Gaji.Value = IIf(Not IsNull(!gaji), !gaji, Null)
        
        Txt_Ket.Text = IIf(Not IsNull(!ket), !ket, "")
        TDB_Uang_Makan.Value = IIf(Not IsNull(!Uang_Makan), !Uang_Makan, Null)
        TDB_Tunjangan.Value = IIf(Not IsNull(!Tunjangan), !Tunjangan, Null)
        
        Dim fl_gaji As String
            fl_gaji = IIf(Not IsNull(!flag_gaji), !flag_gaji, "")
            
            If fl_gaji = "b" Then
                Opt_Bulan.Value = True
            ElseIf fl_gaji = "h" Then
                Opt_Hari.Value = True
            End If
        
        TDB_Tgl_Masuk.Value = IIf(Not IsNull(!Tgl_Masuk), !Tgl_Masuk, Date)
        
        Lbl_Kode_Counter.Caption = IIf(Not IsNull(!kode_counter), !kode_counter, "")
        Lbl_Nama_Counter.Caption = IIf(Not IsNull(!nama_counter), !nama_counter, "")
        
        If .RecordCount = 0 Then
            Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
        Else
            Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
        End If
    End With
    
End Sub

Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then TDB_Rubah.Visible = False: Cmd_Batal_Click
End Sub

Private Sub Opt_Aka_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Opt_Bulan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tunjangan.SetFocus
End Sub

Private Sub Opt_Hari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tunjangan.SetFocus
End Sub

Private Sub Opt_Istana_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub TDB_Counter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Counter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Counter.Top = TDB_Counter.Top - (yold - Y)
   TDB_Counter.Left = TDB_Counter.Left - (xold - X)
End If

End Sub

Private Sub TDB_Counter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Gaji_GotFocus()
    Call Focus_(TDB_Gaji)
End Sub

Private Sub TDB_Gaji_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Opt_Bulan.SetFocus
    End If
End Sub

Private Sub TDB_Gaji_LostFocus()
    
    If TDB_Gaji.ValueIsNull Then
        TDB_Gaji.Value = Null
    End If
    
End Sub

Private Sub TDB_Hutang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub TDB_Rubah_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Rubah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Rubah.Top = TDB_Rubah.Top - (yold - Y)
   TDB_Rubah.Left = TDB_Rubah.Left - (xold - X)
End If

End Sub

Private Sub TDB_Rubah_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Hapus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Hapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Hapus.Top = TDB_Hapus.Top - (yold - Y)
   TDB_Hapus.Left = TDB_Hapus.Left - (xold - X)
End If

End Sub

Private Sub TDB_Hapus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Daftar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Daftar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Daftar.Top = TDB_Daftar.Top - (yold - Y)
   TDB_Daftar.Left = TDB_Daftar.Left - (xold - X)
End If

End Sub

Private Sub TDB_Daftar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Tgl_Lhr_GotFocus()
    Call Focus_(TDB_Tgl_Lhr)
End Sub

Private Sub TDB_Tgl_Lhr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Cbo_Jenis_Kelamin_Click
        Cbo_Jenis_Kelamin.SetFocus
    End If
End Sub

Private Sub TDB_Tgl_Lhr_LostFocus()

On Error GoTo err_handler

    Dim Tahun As Long
        
        Tahun = Year(Now) - Year(TDB_Tgl_Lhr.Text)
        Lbl_Umur.Caption = Tahun & " Thn"

On Error GoTo 0
Exit Sub

err_handler:
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
        
End Sub

Private Sub TDB_Tgl_Masuk_GotFocus()
    Call Focus_(TDB_Tgl_Masuk)
End Sub

Private Sub TDB_Tgl_Masuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub TDB_Tunjangan_GotFocus()
    Call Focus_(TDB_Tunjangan)
End Sub

Private Sub TDB_Tunjangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Uang_Makan.SetFocus
End Sub

Private Sub TDB_Tunjangan_LostFocus()
    If TDB_Tunjangan.ValueIsNull Then
        TDB_Tunjangan.Value = Null
    End If
End Sub

Private Sub TDB_Uang_Makan_GotFocus()
    Call Focus_(TDB_Uang_Makan)
End Sub

Private Sub TDB_Uang_Makan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Ket.SetFocus
End Sub

Private Sub TDB_Uang_Makan_LostFocus()
    If TDB_Uang_Makan.ValueIsNull Then
        TDB_Uang_Makan.Value = Null
    End If
End Sub

Private Sub Txt_Alamat_1_GotFocus()
    Call Focus_(Txt_Alamat_1)
End Sub

Private Sub Txt_Alamat_1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Alamat_2.SetFocus
    End If
End Sub

Private Sub Txt_Alamat_2_GotFocus()
    Call Focus_(Txt_Alamat_2)
End Sub

Private Sub Txt_Alamat_2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Alamat_3.SetFocus
    End If
End Sub

Private Sub Txt_Alamat_3_GotFocus()
    Call Focus_(Txt_Alamat_3)
End Sub

Private Sub Txt_Alamat_3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Kodepos.SetFocus
    End If
End Sub

Private Sub Txt_Cr_Counter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Counter.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Counter.Visible = False: Cmd_Browse_Counter.SetFocus
End Sub

Private Sub Txt_Cr_Counter_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select top 100 * from Tb_Mast_Counter"
        
        Select Case Index
            Case 0
                sql = sql & " where Kode like '%" & Trim(Txt_Cr_Counter(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Counter like '%" & Trim(Txt_Cr_Counter(1).Text) & "%'"
        End Select
        
        sql = sql & " order by Kode,Nama_Counter asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
            
            Set Grid_Counter.DataSource = rs
                Grid_Counter.Refresh
        
End Sub

Private Sub Txt_Cr_Daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
           
    Dim sql As String
        sql = "select top 100 * from VIEW_Karyawan"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 2

End Sub

Private Sub Txt_Cr_Hapus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Hapus_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
        sql = "select top 100 * from VIEW_Karyawan"
            
    If Txt_Cr_Hapus(0).Text <> "" Or Txt_Cr_Hapus(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Hapus(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Hapus(1).Text) & "%'"
        End Select
    End If

    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 1
    
End Sub

Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)


            
    Dim sql As String
        sql = "select top 100 * from View_Karyawan"
        
    If Txt_Cr_Rubah(0).Text <> "" Or Txt_Cr_Rubah(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where Kode_Karyawan like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
            Case 1
                sql = sql & " where Nama_Karyawan like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by Kode_Karyawan asc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Isi_grid_transaksi Rs_Nav, 0
    
End Sub


Private Sub Txt_Ket_GotFocus()
    Call Focus_(Txt_Ket)
End Sub

Private Sub Txt_Ket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Browse_Counter.SetFocus
End Sub

Private Sub Txt_Kode_Agama_Change()
    With Txt_Kode_Agama
        If .Text = "01" Then
            Cbo_Agama.ListIndex = 0
        ElseIf .Text = "02" Then
            Cbo_Agama.ListIndex = 1
        ElseIf .Text = "03" Then
            Cbo_Agama.ListIndex = 2
        ElseIf .Text = "04" Then
            Cbo_Agama.ListIndex = 3
        ElseIf .Text = "05" Then
            Cbo_Agama.ListIndex = 4
        End If
    End With
End Sub

Private Sub Txt_Kode_Jenis_Kelamin_Change()
    With Txt_Kode_Jenis_Kelamin
        If .Text = "01" Then
            Cbo_Jenis_Kelamin.ListIndex = 0
        Else
            Cbo_Jenis_Kelamin.ListIndex = 1
        End If
    End With
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_handler
    
    Dim n As Object
    If KeyCode = 13 And Txt_Kode.Text <> "" Then
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select Kode_Karyawan from Tb_Karyawan where Kode_Karyawan='" & Trim(Txt_Kode.Text) & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
     If Not rs.EOF Then
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Kode Sudah ada ...", vbOKOnly + vbInformation, "Informasi"))

    For Each n In Me
    
        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") And UCase(n.Name) <> UCase("txt_kode") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is TDBDate Then n.Enabled = False
        If TypeOf n Is TDBNumber Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

    Set n = Nothing
                
                Txt_Kode.SetFocus
                Cmd_Simpan.Enabled = False
                 
                On Error GoTo 0
                Exit Sub
        Else
                    
                    For Each n In Me
                    
                        If TypeOf n Is TextBox Then
                        
                         If Not (UCase(n.Name) = UCase("Txt_Kode") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan") Then
                            n.Enabled = True
                         End If
                         
                         If Not (Left(UCase(n.Name), 9) = UCase("Txt_Kode")) Then
                            n.Text = ""
                         End If
                         
                        End If
                        
                       If TypeOf n Is TDBDate Then n.Enabled = True
                        If TypeOf n Is TDBNumber Then
                            n.Enabled = True
                            n.Text = ""
                        End If
                        If TypeOf n Is ComboBox Then n.Enabled = True
                        If TypeOf n Is OptionButton Then n.Enabled = True
                        If TypeOf n Is CommandButton Then
                            If n.Caption = "..." Then
                                n.Enabled = True
                            End If
                        End If
                 
                        
                    Next
                    
                    Set n = Nothing

                Txt_Ket.Text = ""
                Txt_Ket.Enabled = True
                Lbl_Umur.Caption = "0 Thn"
                Lbl_Nama_Counter.Caption = ""
                Lbl_Kode_Counter.Caption = ""
                Cmd_Simpan.Enabled = True
                Txt_Nama.SetFocus
                
        End If
            
    End If
    
On Error GoTo 0
Exit Sub

err_handler:
    
    Dim p As Integer
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear

End Sub

Private Sub Txt_Kodepos_GotFocus()
    Call Focus_(Txt_Kodepos)
End Sub

Private Sub Txt_Kodepos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp_Rumah.SetFocus
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Txt_Tempat_Lhr.SetFocus
    End If
End Sub

Private Sub Txt_Telp_Hp_GotFocus()
    Call Focus_(Txt_Telp_Hp)
End Sub

Private Sub Txt_Telp_Hp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TDB_Tgl_Masuk.SetFocus
End Sub

Private Sub Txt_Telp_Rumah_GotFocus()
    Call Focus_(Txt_Telp_Rumah)
End Sub

Private Sub Txt_Telp_Rumah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp_Hp.SetFocus
End Sub

Private Sub Txt_Tempat_Lhr_GotFocus()
    Call Focus_(Txt_Tempat_Lhr)
End Sub

Private Sub Txt_Tempat_Lhr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TDB_Tgl_Lhr.SetFocus
    End If
End Sub
