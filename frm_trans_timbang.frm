VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_trans_timbang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transaksi Penimbangan"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_trans_timbang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   4095
      Left            =   -7560
      TabIndex        =   89
      Top             =   2520
      Visible         =   0   'False
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   7223
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_trans_timbang.frx":08CA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_trans_timbang.frx":08E6
      Childs          =   "frm_trans_timbang.frx":0992
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   315
         Index           =   2
         Left            =   5760
         TabIndex        =   98
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   97
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   96
         Top             =   600
         Width           =   1455
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
         Index           =   2
         Left            =   240
         TabIndex        =   90
         Top             =   360
         Width           =   7455
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   3015
         Left            =   240
         OleObjectBlob   =   "frm_trans_timbang.frx":09AE
         TabIndex        =   91
         Top             =   960
         Width           =   7455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   95
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   33
         Left            =   2880
         TabIndex        =   94
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   93
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Ambil"
         Height          =   195
         Index           =   32
         Left            =   5040
         TabIndex        =   92
         Top             =   600
         Width           =   630
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   4095
      Left            =   -7680
      TabIndex        =   78
      Top             =   3240
      Visible         =   0   'False
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   7223
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_trans_timbang.frx":51DE
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_trans_timbang.frx":51FA
      Childs          =   "frm_trans_timbang.frx":52A6
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   2
         Left            =   5880
         TabIndex        =   86
         Top             =   600
         Width           =   1455
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
         Index           =   1
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   7455
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   80
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   79
         Top             =   600
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   3015
         Left            =   240
         OleObjectBlob   =   "frm_trans_timbang.frx":52C2
         TabIndex        =   82
         Top             =   960
         Width           =   7455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Ambil"
         Height          =   195
         Index           =   31
         Left            =   5040
         TabIndex        =   87
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   85
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   30
         Left            =   2880
         TabIndex        =   84
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   29
         Left            =   240
         TabIndex        =   83
         Top             =   600
         Width           =   855
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Anggota 
      Height          =   3735
      Left            =   -6000
      TabIndex        =   68
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
      MouseIcon       =   "frm_trans_timbang.frx":9AF1
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_trans_timbang.frx":9B0D
      Childs          =   "frm_trans_timbang.frx":9BB9
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
         TabIndex        =   71
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox Txt_Cr_Anggota 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   70
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Txt_Cr_Anggota 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   69
         Top             =   600
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Anggota 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "frm_trans_timbang.frx":9BD5
         TabIndex        =   72
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   75
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   37
         Left            =   2880
         TabIndex        =   74
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   73
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.TextBox tno_anggota 
      Height          =   300
      Left            =   2040
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8625
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox ck_wjb 
         Alignment       =   1  'Right Justify
         Caption         =   "I"
         Height          =   195
         Left            =   2040
         TabIndex        =   67
         Top             =   5310
         Width           =   255
      End
      Begin VB.CommandButton cbr_anggota 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4050
         TabIndex        =   66
         Top             =   335
         Width           =   375
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
         Left            =   4440
         TabIndex        =   59
         Top             =   7680
         Width           =   3615
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   2640
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Daftar"
            Height          =   495
            Left            =   1800
            TabIndex        =   61
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   495
            Left            =   960
            TabIndex        =   62
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   495
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   495
            Left            =   960
            TabIndex        =   65
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "&Cetak Slip Bukti"
         Height          =   735
         Left            =   8160
         TabIndex        =   58
         Top             =   7800
         Width           =   975
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
         Left            =   240
         TabIndex        =   53
         Top             =   7680
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
            TabIndex        =   54
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
            TabIndex        =   55
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
            TabIndex        =   56
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
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox t_k_brt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox t_b_brt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox t_kp_susut 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t_k_susut 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox t_bp_susut 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t_b_susut 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   17
         Text            =   "0"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox t_k_timbang 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   15
         Text            =   "0"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox t_b_Timbang 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   13
         Text            =   "0"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox tnama_anggota 
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   300
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   50987009
         CurrentDate     =   39372
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_krm 
         Height          =   300
         Left            =   2040
         TabIndex        =   45
         Top             =   4200
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":CB5A
         Caption         =   "frm_trans_timbang.frx":CB7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":CBE6
         Keys            =   "frm_trans_timbang.frx":CC04
         Spin            =   "frm_trans_timbang.frx":CC4E
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_adm 
         Height          =   300
         Left            =   2040
         TabIndex        =   46
         Top             =   4560
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":CC76
         Caption         =   "frm_trans_timbang.frx":CC96
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":CD02
         Keys            =   "frm_trans_timbang.frx":CD20
         Spin            =   "frm_trans_timbang.frx":CD6A
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_wjb 
         Height          =   300
         Left            =   2400
         TabIndex        =   47
         Top             =   5280
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":CD92
         Caption         =   "frm_trans_timbang.frx":CDB2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":CE1E
         Keys            =   "frm_trans_timbang.frx":CE3C
         Spin            =   "frm_trans_timbang.frx":CE86
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_skr 
         Height          =   300
         Left            =   2040
         TabIndex        =   48
         Top             =   5640
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":CEAE
         Caption         =   "frm_trans_timbang.frx":CECE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":CF3A
         Keys            =   "frm_trans_timbang.frx":CF58
         Spin            =   "frm_trans_timbang.frx":CFA2
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
      Begin TDBNumber6Ctl.TDBNumber tdb_cc_ppk 
         Height          =   300
         Left            =   2040
         TabIndex        =   49
         Top             =   6000
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":CFCA
         Caption         =   "frm_trans_timbang.frx":CFEA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D056
         Keys            =   "frm_trans_timbang.frx":D074
         Spin            =   "frm_trans_timbang.frx":D0BE
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
      Begin TDBNumber6Ctl.TDBNumber tdb_cc_krd 
         Height          =   300
         Left            =   2040
         TabIndex        =   50
         Top             =   6360
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D0E6
         Caption         =   "frm_trans_timbang.frx":D106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D172
         Keys            =   "frm_trans_timbang.frx":D190
         Spin            =   "frm_trans_timbang.frx":D1DA
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
      Begin TDBNumber6Ctl.TDBNumber tdb_hrg_br 
         Height          =   300
         Left            =   2040
         TabIndex        =   51
         Top             =   4920
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D202
         Caption         =   "frm_trans_timbang.frx":D222
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D28E
         Keys            =   "frm_trans_timbang.frx":D2AC
         Spin            =   "frm_trans_timbang.frx":D2F6
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_jtr 
         Height          =   300
         Left            =   2040
         TabIndex        =   52
         Top             =   7080
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D31E
         Caption         =   "frm_trans_timbang.frx":D33E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D3AA
         Keys            =   "frm_trans_timbang.frx":D3C8
         Spin            =   "frm_trans_timbang.frx":D412
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,###;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,###"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999999
         MinValue        =   -9999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber t_b_hrg 
         Height          =   300
         Left            =   4320
         TabIndex        =   76
         Top             =   3240
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D43A
         Caption         =   "frm_trans_timbang.frx":D45A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D4C6
         Keys            =   "frm_trans_timbang.frx":D4E4
         Spin            =   "frm_trans_timbang.frx":D52E
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber t_k_hrg 
         Height          =   300
         Left            =   2040
         TabIndex        =   77
         Top             =   3240
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D556
         Caption         =   "frm_trans_timbang.frx":D576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D5E2
         Keys            =   "frm_trans_timbang.frx":D600
         Spin            =   "frm_trans_timbang.frx":D64A
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_adm1 
         Height          =   300
         Left            =   4320
         TabIndex        =   99
         Top             =   4560
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D672
         Caption         =   "frm_trans_timbang.frx":D692
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D6FE
         Keys            =   "frm_trans_timbang.frx":D71C
         Spin            =   "frm_trans_timbang.frx":D766
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_krm1 
         Height          =   300
         Left            =   4320
         TabIndex        =   100
         Top             =   4200
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D78E
         Caption         =   "frm_trans_timbang.frx":D7AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D81A
         Keys            =   "frm_trans_timbang.frx":D838
         Spin            =   "frm_trans_timbang.frx":D882
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_hrg_br1 
         Height          =   300
         Left            =   4320
         TabIndex        =   101
         Top             =   4920
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D8AA
         Caption         =   "frm_trans_timbang.frx":D8CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":D936
         Keys            =   "frm_trans_timbang.frx":D954
         Spin            =   "frm_trans_timbang.frx":D99E
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
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1028849669
         MinValueVT      =   1598423045
      End
      Begin TDBNumber6Ctl.TDBNumber tdb_dp 
         Height          =   300
         Left            =   2040
         TabIndex        =   104
         Top             =   6720
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         Calculator      =   "frm_trans_timbang.frx":D9C6
         Caption         =   "frm_trans_timbang.frx":D9E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_trans_timbang.frx":DA52
         Keys            =   "frm_trans_timbang.frx":DA70
         Spin            =   "frm_trans_timbang.frx":DABA
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP :"
         Height          =   195
         Index           =   35
         Left            =   1680
         TabIndex        =   105
         Top             =   6720
         Width           =   300
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8400
         TabIndex        =   88
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Diterima :"
         Height          =   195
         Index           =   28
         Left            =   720
         TabIndex        =   44
         Top             =   7080
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Bersih /Kg :"
         Height          =   195
         Index           =   25
         Left            =   675
         TabIndex        =   43
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cicilan Kredit :"
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   42
         Top             =   6360
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cicilan Pupuk :"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   41
         Top             =   6000
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simpanan Sukarela :"
         Height          =   195
         Index           =   24
         Left            =   480
         TabIndex        =   40
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simpanan Wajib :"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   39
         Top             =   5280
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiriman /Kg :"
         Height          =   195
         Index           =   22
         Left            =   795
         TabIndex        =   38
         Top             =   4200
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrasi /Kg :"
         Height          =   195
         Index           =   21
         Left            =   675
         TabIndex        =   37
         Top             =   4560
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BIAYA && JUMLAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   2040
         TabIndex        =   36
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIMBANGAN && HARGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   35
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   20
         Left            =   3480
         TabIndex        =   34
         Top             =   3240
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   19
         Left            =   5760
         TabIndex        =   33
         Top             =   3240
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   18
         Left            =   3480
         TabIndex        =   32
         Top             =   2880
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   17
         Left            =   5760
         TabIndex        =   30
         Top             =   2880
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   4200
         TabIndex        =   28
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   2970
         TabIndex        =   27
         Top             =   2520
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   14
         Left            =   3960
         TabIndex        =   26
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   13
         Left            =   2760
         TabIndex        =   24
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   6480
         TabIndex        =   22
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   5235
         TabIndex        =   21
         Top             =   2520
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   10
         Left            =   6240
         TabIndex        =   20
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   9
         Left            =   5040
         TabIndex        =   18
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   8
         Left            =   3480
         TabIndex        =   16
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   7
         Left            =   5760
         TabIndex        =   14
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Pabrik /Kg :"
         Height          =   195
         Index           =   6
         Left            =   705
         TabIndex        =   9
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Bersih/Timb Pabrik :"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   2880
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penyusutan :"
         Height          =   195
         Index           =   4
         Left            =   1020
         TabIndex        =   7
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Ditimbang :"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   6
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BASAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   4320
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KERING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Pengambilan :"
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   3
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   2
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota :"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KERING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   8
         Left            =   2040
         TabIndex        =   103
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BASAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   7
         Left            =   4320
         TabIndex        =   102
         Top             =   3840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_trans_timbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long
Dim simp_wjb As Double
Dim idrubah As Long
Dim simp_wajib, simp_skr As Double
Dim krimawal, krimawal1 As Double

Private Sub cari_biaya_adm_perkg()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Biaya_Adm"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                tdb_adm.Value = IIf(Not IsNull(!harga_kering), !harga_kering, Null)
                tdb_adm1.Value = IIf(Not IsNull(!harga_basah), !harga_basah, Null)
            Else
                tdb_adm.Value = Null
                tdb_adm1.Value = Null
            End If
        End With
        
    
End Sub

Private Sub cari_biaya_kirim_perkg()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Biaya_Kirim"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                krimawal = IIf(Not IsNull(!harga_kering), !harga_kering, 0)
                krimawal1 = IIf(Not IsNull(!harga_basah), !harga_basah, 0)
            Else
                krimawal = 0
                krimawal1 = 0
            End If
        End With
        
        tdb_krm.Value = IIf((krimawal = 0), Null, krimawal)
        tdb_krm1.Value = IIf((krimawal1 = 0), Null, krimawal1)
        
    
End Sub

Private Sub cari_penyusutan()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Penyusutan"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            If Not .EOF Then
                t_k_susut.Text = IIf(Not IsNull(!kering), !kering, 0)
                t_b_susut.Text = IIf(Not IsNull(!basah), !basah, 0)
            Else
                t_k_susut.Text = 0
                t_b_susut.Text = 0
            End If
        End With
        
    
End Sub

Private Sub isi_semua(ByVal rec As Recordset)
    
    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        idrubah = !id
        tno_anggota.Text = IIf(Not IsNull(!no_anggota), !no_anggota, "")
        tnama_anggota.Text = IIf(Not IsNull(!nama), !nama, "")
        dtp_tgl.Value = IIf(Not IsNull(!Tgl_Ambil), !Tgl_Ambil, Date)
        t_k_timbang.Text = IIf(Not IsNull(!Berat_K), !Berat_K, 0)
        t_b_Timbang.Text = IIf(Not IsNull(!Berat_B), !Berat_B, 0)
        
        t_k_susut.Text = IIf(Not IsNull(!Penyusutan_KP), !Penyusutan_KP, 0)
        t_b_susut.Text = IIf(Not IsNull(!Penyusutan_BP), !Penyusutan_BP, 0)
        
        t_kp_susut.Text = IIf(Not IsNull(!Penyusutan_K), !Penyusutan_K, 0)
        t_bp_susut.Text = IIf(Not IsNull(!Penyusutan_B), !Penyusutan_B, 0)
        
        t_k_brt.Text = IIf(Not IsNull(!Berat_B_K), !Berat_B_K, 0)
        t_b_brt.Text = IIf(Not IsNull(!Berat_B_B), !Berat_B_B, 0)
        t_k_hrg.Value = IIf(Not IsNull(!Harga_K), !Harga_K, Null)
        t_b_hrg.Value = IIf(Not IsNull(!Harga_B), !Harga_B, Null)
        tdb_adm.Value = IIf(Not IsNull(!By_Admin_K), !By_Admin_K, Null)
        tdb_adm1.Value = IIf(Not IsNull(!By_Admin_B), !By_Admin_B, Null)
        tdb_krm.Value = IIf(Not IsNull(!By_Kirim_K), !By_Kirim_K, Null)
        tdb_krm1.Value = IIf(Not IsNull(!By_Kirim_B), !By_Kirim_B, Null)
        tdb_wjb.Value = IIf(Not IsNull(!Sm_Wajib), !Sm_Wajib, Null)
        
        simp_wajib = IIf(Not IsNull(!Sm_Wajib), !Sm_Wajib, Null)
        
        If tdb_wjb.ValueIsNull Then
            ck_wjb.Value = vbUnchecked
        Else
            ck_wjb.Value = vbChecked
        End If
        
        tdb_skr.Value = IIf(Not IsNull(!Sm_Sukarela), !Sm_Sukarela, Null)
        
        simp_skr = IIf(Not IsNull(!Sm_Sukarela), !Sm_Sukarela, Null)
        
        tdb_cc_ppk.Value = IIf(Not IsNull(!Cl_Pupuk), !Cl_Pupuk, Null)
        tdb_cc_krd.Value = IIf(Not IsNull(!Cl_Kredit), !Cl_Kredit, Null)
        tdb_dp.Value = IIf(Not IsNull(!dp), !dp, Null)
        tdb_hrg_br.Value = IIf(Not IsNull(!Hrg_Bersih_K), !Hrg_Bersih_K, Null)
        tdb_hrg_br1.Value = IIf(Not IsNull(!Hrg_Bersih_B), !Hrg_Bersih_B, Null)
        tdb_jtr.Value = IIf(Not IsNull(!jml_trm), !jml_trm, Null)
        
        
            If .RecordCount = 0 Then
                Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
            Else
                Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
            End If
            
    End With
    
End Sub

Private Sub total_harga()
'    On Error Resume Next
    
    Dim brt_tb_bsh As Double
    Dim brt_tb_krg As Double
        
        If t_b_Timbang.Text = "" Then
            brt_tb_bsh = 0
        Else
            brt_tb_bsh = Replace(t_b_Timbang.Text, ".", ",")
        End If
        
        If t_k_timbang.Text = "" Then
            brt_tb_krg = 0
        Else
            brt_tb_krg = Replace(t_k_timbang.Text, ".", ",")
        End If
        
    Dim sst_bsh As String
    Dim sst_krg As String
        
        If t_bp_susut.Text = "" Then
            sst_bsh = 0
        Else
            sst_bsh = t_bp_susut.Text
            
            If sst_bsh <> 0 Then sst_bsh = Replace(sst_bsh, ".", ",")
            
        End If
        
        If t_kp_susut.Text = "" Then
            sst_krg = 0
        Else
            sst_krg = t_kp_susut.Text
'
            If sst_krg <> 0 Then sst_krg = Replace(sst_krg, ".", ",")
            
        End If
        
    
    Dim jmlkrim As Double
        If tdb_krm.ValueIsNull Then
            jmlkrim = 0
        Else
            jmlkrim = Replace(Trim(tdb_krm.Value), ",", "")
        End If
        
        
    Dim totkrm As Double
        totkrm = CDbl(jmlkrim) * CDbl(brt_tb_krg)
    
    Dim jmlkrim1 As Double
        If tdb_krm1.ValueIsNull Then
            jmlkrim1 = 0
        Else
            jmlkrim1 = Replace(Trim(tdb_krm1.Value), ",", "")
        End If
    
    Dim totkrm1 As Double
        totkrm1 = CDbl(jmlkrim1) * CDbl(brt_tb_bsh)

    Dim brt_bsh As Double
    Dim brt_krg As Double
        
        brt_bsh = brt_tb_bsh - CDbl(sst_bsh)
        brt_krg = brt_tb_krg - CDbl(sst_krg)
    
        t_b_brt.Text = FormatNumber(brt_bsh, 2)
        t_k_brt.Text = FormatNumber(brt_krg, 2)


    Dim jmladm As Double
        If tdb_adm.ValueIsNull Then
            jmladm = 0
        Else
            jmladm = Replace(Trim(tdb_adm.Value), ",", "")
        End If
    
    Dim jmladm1 As Double
        If tdb_adm1.ValueIsNull Then
            jmladm1 = 0
        Else
            jmladm1 = Replace(Trim(tdb_adm1.Value), ",", "")
        End If
    
    
    Dim hrg_pk_bsh As Double
        If t_b_hrg.ValueIsNull Then
            hrg_pk_bsh = 0
        Else
            hrg_pk_bsh = Replace(Trim(t_b_hrg.Value), ",", "")
        End If
    
    Dim hrg_pk_krg As Double
        If t_k_hrg.ValueIsNull Then
            hrg_pk_krg = 0
        Else
            hrg_pk_krg = Replace(Trim(t_k_hrg.Value), ",", "")
        End If
    
    
    Dim jmlwjb As Double
        If tdb_wjb.ValueIsNull Then
            jmlwjb = 0
        Else
            jmlwjb = Replace(Trim(tdb_wjb.Value), ",", "")
        End If
    
    Dim jmlskr As Double
        If tdb_skr.ValueIsNull Then
            jmlskr = 0
        Else
            jmlskr = Replace(Trim(tdb_skr.Value), ",", "")
        End If
    
    Dim jmlcippk As Double
        If tdb_cc_ppk.ValueIsNull Then
            jmlcippk = 0
        Else
            jmlcippk = Replace(Trim(tdb_cc_ppk.Value), ",", "")
        End If
    
    Dim jmlcikrd As Double
        If tdb_cc_krd.ValueIsNull Then
            jmlcikrd = 0
        Else
            jmlcikrd = Replace(Trim(tdb_cc_krd.Value), ",", "")
        End If
    
    Dim dp As Double
        If tdb_dp.ValueIsNull Then
            dp = 0
        Else
            dp = Replace(Trim(tdb_dp.Value), ",", "")
        End If
    
    Dim tot_terima As Double
    Dim tot_terima_krg, tot_terima_bsh As Double
        tot_terima_krg = 0
        tot_terima_bsh = 0
    
    If brt_tb_bsh = 0 And brt_tb_krg <> 0 Then
        tot_terima = ((hrg_pk_krg * brt_krg) - (((brt_tb_krg * jmlkrim) + (brt_krg * jmladm))))
        tot_terima_krg = tot_terima
        
        tot_terima = tot_terima - (jmlwjb + jmlskr + jmlcippk + jmlcikrd + dp)
        
    ElseIf brt_tb_bsh <> 0 And brt_tb_krg = 0 Then
        tot_terima = ((hrg_pk_bsh * brt_bsh) - (((brt_tb_bsh * jmlkrim1) + (brt_bsh * jmladm1))))
        tot_terima_bsh = tot_terima
        
        tot_terima = tot_terima - (jmlwjb + jmlskr + jmlcippk + jmlcikrd + dp)
        
    ElseIf brt_tb_bsh <> 0 And brt_tb_krg <> 0 Then
        tot_terima = ((hrg_pk_krg * brt_krg) - (((brt_tb_krg * jmlkrim) + (brt_krg * jmladm))))
        
        tot_terima_krg = tot_terima
        
        tot_terima_bsh = ((hrg_pk_bsh * brt_bsh) - (((brt_tb_bsh * jmlkrim1) + (brt_bsh * jmladm1))))
        
        tot_terima = tot_terima + ((hrg_pk_bsh * brt_bsh) - (((brt_tb_bsh * jmlkrim1) + (brt_bsh * jmladm1))))
        
        tot_terima = tot_terima - (jmlwjb + jmlskr + jmlcippk + jmlcikrd + dp)
        
    End If
    
    If tot_terima = 0 Then
        tdb_jtr.Value = Null
    Else
        tdb_jtr.Value = tot_terima
    End If
    
    Dim pers_k As Double
        If t_k_susut.Text = "" Then
            pers_k = 0
        Else
            pers_k = Trim(t_k_susut.Text)
        End If
    
    Dim pers_b As Double
        If t_b_susut.Text = "" Then
            pers_b = 0
        Else
            pers_b = Trim(t_b_susut.Text)
        End If
    
    Dim nilai_s_k As Double
        nilai_s_k = hrg_pk_krg * (pers_k / 100)
    
    Dim nilai_s_b As Double
        nilai_s_b = hrg_pk_bsh * (pers_b / 100)
    
        Dim bersih_k, bersih_b As Double
        If tot_terima <> 0 And brt_tb_krg <> 0 Then
            bersih_k = tot_terima_krg / brt_tb_krg '- (nilai_s_k + jmladm + jmlkrim)        '(hrg_pk_krg * brt_krg) - ((brt_tb_krg * jmlkrim) + (brt_krg * jmladm))
        End If
        
        If tot_terima <> 0 And brt_tb_bsh <> 0 Then
            bersih_b = tot_terima_bsh / brt_tb_bsh '- (nilai_s_b + jmladm1 + jmlkrim1)   '(hrg_pk_bsh * brt_bsh) - ((brt_tb_bsh * jmlkrim1) + (brt_bsh * jmladm1))
        End If
        
'        Dim bersih_k1 As Double
'            bersih_k1 = hrg_pk_krg - (jmlkrim + jmladm)
'
'        Dim bersih_b1 As Double
'            bersih_b1 = hrg_pk_bsh - (jmlkrim1 + jmladm1)
            
        If bersih_k = 0 Then
            tdb_hrg_br.Value = Null
        Else
            tdb_hrg_br.Value = CCur(bersih_k)
        End If
    
        If bersih_b = 0 Then
            tdb_hrg_br1.Value = Null
        Else
            tdb_hrg_br1.Value = CCur(bersih_b)
        End If

End Sub

Private Sub cek_harga_perkg()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select harga_kering,harga_basah from Tb_Harga_Karet"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            t_b_hrg.Value = IIf(Not IsNull(!harga_kering), !harga_kering, Null)
            t_k_hrg.Value = IIf(Not IsNull(!harga_basah), !harga_basah, Null)
        Else
            t_b_hrg.Value = Null
            t_k_hrg.Value = Null
        End If
    End With
    
    
End Sub

Private Sub cek_besar_simp_wajib()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select jumlah from Tb_Simpanan_Wajib"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            simp_wjb = IIf(Not IsNull(!jumlah), !jumlah, 0)
        Else
            simp_wjb = 0
        End If
    End With
    
End Sub

Private Sub cbr_anggota_Click()

    With TDB_Anggota
        
        If .Visible = False Then
        
        .Left = Picture1.Left + cbr_anggota.Left + cbr_anggota.Width / 2 - .Width / 2
        .Top = Picture1.Top + cbr_anggota.Top + cbr_anggota.Height + 15
        
        Txt_Cr_Anggota(0).Text = ""
        Txt_Cr_Anggota(1).Text = ""
        
        Txt_Cr_Anggota_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Anggota(0).SetFocus
        
        Else
            .Visible = False
        End If
        
    End With


End Sub

Private Sub ck_wjb_Click()
    
    If ck_wjb.Value = vbChecked Then
        If simp_wjb = 0 Then
            tdb_wjb.Value = Null
        Else
            tdb_wjb.Value = simp_wjb
        End If
    Else
        tdb_wjb.Value = Null
    End If
    
End Sub

Private Sub ck_wjb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_skr.SetFocus
End Sub

Private Sub Cmd_Batal_Click()

    Frame_Nav.Enabled = True
    rubah = False
             
        Cmd_Tambah.Visible = True
        
        Cmd_Tambah.Enabled = True
    
        Cmd_Simpan.Visible = False
        Cmd_Rubah.Visible = True
        Cmd_Rubah.Enabled = True
'        Cmd_Hapus.Enabled = True
        Cmd_Daftar.Enabled = True
        Cmd_Keluar.Enabled = True
        
        cmd_cetak.Enabled = True
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is DTPicker Then n.Enabled = False
        If TypeOf n Is CheckBox Then n.Enabled = False
        
        If TypeOf n Is TDBNumber Then
            n.Enabled = False
        End If
        
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
        If TypeOf n Is OptionButton Then n.Enabled = False

Next

Set n = Nothing

Txt_Cr_Daftar_KeyUp 0, 0, 0

Cmd_Navigasi_Click 0

End Sub

Private Sub cmd_cetak_Click()

    If tno_anggota.Text = "" Then Exit Sub
    If idrubah = Empty Then Exit Sub

    
    Mysq = "select * from view_Penimbangan where id=" & idrubah
    
    frm_lap_bukt_timbang.Show

End Sub

Private Sub Cmd_Daftar_Click()

Frame_Nav.Enabled = False
With TDB_Daftar

If .Visible = False Then
    
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
    
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
'    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_cetak.Enabled = False
    
    Txt_Cr_Daftar(0).Text = ""
    Txt_Cr_Daftar(1).Text = ""
    Txt_Cr_Daftar(2).Text = ""
    
    Txt_Cr_Daftar_KeyUp 0, 0, 0
    
    .Visible = True
    
    Txt_Cr_Daftar(0).SetFocus
    
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
    
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
'    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_cetak.Enabled = False
    
    Txt_Cr_Rubah(0).Text = ""
    Txt_Cr_Rubah(1).Text = ""
    Txt_Cr_Rubah(2).Text = ""
    
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

    Dim jmlkrim As Double
        If tdb_krm.ValueIsNull Then
            jmlkrim = 0
        Else
            jmlkrim = Replace(Trim(tdb_krm.Value), ",", "")
        End If
    
    Dim jmlkrim1 As Double
        If tdb_krm1.ValueIsNull Then
            jmlkrim1 = 0
        Else
            jmlkrim1 = Replace(Trim(tdb_krm1.Value), ",", "")
        End If
    
    Dim jmladm As Double
        If tdb_adm.ValueIsNull Then
            jmladm = 0
        Else
            jmladm = Replace(Trim(tdb_adm.Value), ",", "")
        End If
    
    Dim jmladm1 As Double
        If tdb_adm1.ValueIsNull Then
            jmladm1 = 0
        Else
            jmladm1 = Replace(Trim(tdb_adm1.Value), ",", "")
        End If
    
    Dim jmlwjb As Double
        If tdb_wjb.ValueIsNull Then
            jmlwjb = 0
        Else
            jmlwjb = Replace(Trim(tdb_wjb.Value), ",", "")
        End If
    
    Dim jmlskr As Double
        If tdb_skr.ValueIsNull Then
            jmlskr = 0
        Else
            jmlskr = Replace(Trim(tdb_skr.Value), ",", "")
        End If
    
    Dim jmlcippk As Double
        If tdb_cc_ppk.ValueIsNull Then
            jmlcippk = 0
        Else
            jmlcippk = Replace(Trim(tdb_cc_ppk.Value), ",", "")
        End If
    
    Dim jmlcikrd As Double
        If tdb_cc_krd.ValueIsNull Then
            jmlcikrd = 0
        Else
            jmlcikrd = Replace(Trim(tdb_cc_krd.Value), ",", "")
        End If

    Dim hrg_pk_bsh As Double
        If t_b_hrg.ValueIsNull Then
            hrg_pk_bsh = 0
        Else
            hrg_pk_bsh = Replace(Trim(t_b_hrg.Value), ",", "")
        End If
    
    Dim hrg_pk_krg As Double
        If t_k_hrg.ValueIsNull Then
            hrg_pk_krg = 0
        Else
            hrg_pk_krg = Replace(Trim(t_k_hrg.Value), ",", "")
        End If
    
    Dim hrg_brs As Double
        If tdb_hrg_br.ValueIsNull Then
            hrg_brs = 0
        Else
            hrg_brs = Replace(Trim(tdb_hrg_br.Value), ",", "")
        End If
    
    Dim hrg_brs1 As Double
        If tdb_hrg_br1.ValueIsNull Then
            hrg_brs1 = 0
        Else
            hrg_brs1 = Replace(Trim(tdb_hrg_br1.Value), ",", "")
        End If
    
    Dim jml_trm As Double
        If tdb_jtr.ValueIsNull Then
            jml_trm = 0
        Else
            jml_trm = Replace(Trim(tdb_jtr.Value), ",", "")
        End If
    
    Dim dp As Double
        If tdb_dp.ValueIsNull Then
            dp = 0
        Else
            dp = Replace(Trim(tdb_dp.Value), ",", "")
        End If
    
    Dim user_active As String
        user_active = Mid(Utama.StatusBar1.Panels(1).Text, 16, Len(Utama.StatusBar1.Panels(1).Text))
        
    
    Dim sql As String
    Dim rs As Recordset
    
    kon.BeginTrans
    
    If rubah = False Then
        
        sql = "insert into Tb_Penimbangan (No_Anggota,Nama,Tgl_Ambil,Berat_K,Berat_B,Penyusutan_KP,Penyusutan_BP,Penyusutan_K,Penyusutan_B,Berat_B_K,Berat_B_B,Harga_K,Harga_B,By_Admin_K,By_Kirim_K,Sm_Wajib,Sm_Sukarela,Cl_Pupuk,Cl_Kredit,Hrg_Bersih_K,Jml_Trm,User_Active,Tgl_Trans,Jam_Trans,By_Admin_B,By_Kirim_B,Hrg_Bersih_B,DP)"
        sql = sql & " values('" & Trim(tno_anggota.Text) & "','" & Trim(tnama_anggota.Text) & "','" & Format(dtp_tgl.Value, "yyyy/mm/dd") & "'," & Replace(Trim(t_k_timbang.Text), ",", ".") & "," & Replace(Trim(t_b_Timbang.Text), ",", ".") & ""
        sql = sql & "," & Replace(Trim(t_k_susut.Text), ",", ".") & "," & Replace(Trim(t_b_susut.Text), ",", ".") & "," & Replace(Trim(t_kp_susut.Text), ",", ".") & "," & Replace(Trim(t_bp_susut.Text), ",", ".") & "," & Replace(Trim(t_k_brt.Text), ",", ".") & "," & Replace(Trim(t_b_brt.Text), ",", ".") & ""
        sql = sql & "," & Replace(Trim(t_k_hrg.Value), ",", "") & "," & Replace(Trim(t_b_hrg.Value), ",", "") & "," & jmladm & "," & jmlkrim & "," & jmlwjb & "," & jmlskr & "," & jmlcippk & "," & jmlcikrd & "," & hrg_brs & "," & jml_trm & ",'" & user_active & "','" & Format(Date, "yyyy/mm/dd") & "','" & Format(Time, "hh:mm:ss") & "'," & jmladm1 & "," & jmlkrim1 & "," & hrg_brs1 & "," & dp & " )"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        tambah_simpanan_wajib_n_sukarela jmlwjb, jmlskr
        
        MsgBox "Data telah disimpan"
        
    Else
        
        sql = "update Tb_Penimbangan set Tgl_Ambil='" & Format(dtp_tgl.Value, "yyyy/mm/dd") & "',Berat_K=" & Replace(Trim(t_k_timbang.Text), ",", ".") & ",Berat_B=" & Replace(Trim(t_b_Timbang.Text), ",", ".") & ",Penyusutan_KP=" & Replace(Trim(t_k_susut.Text), ",", ".") & ",Penyusutan_BP=" & Replace(Trim(t_b_susut.Text), ",", ".") & ""
        sql = sql & ",Penyusutan_K=" & Replace(Trim(t_kp_susut.Text), ",", ".") & ",Penyusutan_B=" & Replace(Trim(t_bp_susut.Text), ",", ".") & ",Berat_B_K=" & Replace(Trim(t_k_brt.Text), ",", ".") & ",Berat_B_B=" & Replace(Trim(t_b_brt.Text), ",", ".") & ",Harga_K=" & Replace(Trim(t_k_hrg.Value), ",", "") & ",Harga_B=" & Replace(Trim(t_b_hrg.Value), ",", "") & ""
        sql = sql & ",By_Admin_K=" & jmladm & ",By_Kirim_K=" & jmlkrim & ",Sm_Wajib=" & jmlwjb & ",Sm_Sukarela=" & jmlskr & ",Cl_Pupuk=" & jmlcippk & ",Cl_Kredit=" & jmlcikrd & ",Hrg_Bersih_K=" & hrg_brs & ",Jml_Trm=" & jml_trm & ",User_Active='" & user_active & "',Tgl_Trans='" & Format(Date, "yyyy/mm/dd") & "',Jam_Trans='" & Format(Time, "hh:mm:ss") & "',"
        sql = sql & "By_Admin_B=" & jmladm1 & ",By_Kirim_B=" & jmlkrim1 & ",Hrg_Bersih_B=" & hrg_brs1 & ",DP=" & dp
        sql = sql & " where id=" & idrubah
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        update_simpanan_wajib_n_sukarela jmlwjb, jmlskr
        
        MsgBox "Data telah dirubah"
        
    End If
    
    kon.CommitTrans
    
    Cmd_Batal_Click
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
        kon.RollbackTrans
        MsgBox Error$
    
End Sub

Private Sub tambah_simpanan_wajib_n_sukarela(ByVal jmlwajib As Double, ByVal jmlsukarela As Double)
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "update tb_anggota set jml_wajib=jml_wajib + " & jmlwajib & ",Jml_Sukarela=jml_sukarela + " & jmlsukarela
        sql = sql & " where no_anggota='" & Trim(tno_anggota.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
    
End Sub

Private Sub update_simpanan_wajib_n_sukarela(ByVal jmlwajib As Double, ByVal jmlsukarela As Double)
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "update tb_anggota set jml_wajib=jml_wajib - " & simp_wajib & ",Jml_Sukarela=jml_sukarela - " & simp_skr
        sql = sql & " where no_anggota='" & Trim(tno_anggota.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        sql = "update tb_anggota set jml_wajib=jml_wajib + " & jmlwajib & ",Jml_Sukarela=jml_sukarela + " & jmlsukarela
        sql = sql & " where no_anggota='" & Trim(tno_anggota.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
  
    
End Sub

Private Sub Cmd_Tambah_Click()

    Frame_Nav.Enabled = False
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
     Cmd_Rubah.Visible = False
     Cmd_Batal.Visible = True
'     Cmd_Hapus.Enabled = False
     Cmd_Daftar.Enabled = False
     Cmd_Keluar.Enabled = False
    
     cmd_cetak.Enabled = False
    
     cbr_anggota.Enabled = True
     tno_anggota.Enabled = True
     tno_anggota.Text = ""
     
     tno_anggota.SetFocus


End Sub

Private Sub dtp_tgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        t_k_timbang.SetFocus
    End If
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
        If TypeOf n Is DTPicker Then n.Enabled = False
        If TypeOf n Is TDBNumber Then n.Enabled = False
        If TypeOf n Is CheckBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

Set n = Nothing

cek_besar_simp_wajib

Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & 0 & " Record"

Txt_Cr_Daftar_KeyUp 0, 0, 0

Cmd_Navigasi_Click 0


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

Private Sub Grid_Anggota_DblClick()
    
    If Grid_Anggota.Row < 0 Then Exit Sub

    
    tno_anggota.Text = Grid_Anggota.Columns(0).Text
     
    tno_anggota_KeyDown 13, 0
    
    TDB_Anggota.Visible = False
    
End Sub

Private Sub Grid_Anggota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Anggota_DblClick
    If KeyCode = vbKeyEscape Then TDB_Anggota.Visible = False: tno_anggota.SetFocus
End Sub

Private Sub grid_daftar_DblClick()
    
    If Grid_Daftar.Row < 0 Then Exit Sub

    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "id='" & Grid_Daftar.Columns(6).Text & "'"

    isi_semua Rs_Nav
    
    cmd_cetak.Enabled = True
    TDB_Daftar.Visible = False
    Frame_Nav.Enabled = True
    Cmd_Navigasi(0).SetFocus
    
End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Grid_Rubah_DblClick()

If Grid_Rubah.Row < 0 Then Exit Sub
    
    Dim nobuk As String
        nobuk = Grid_Rubah.Columns(6).Text
    
    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "id='" & nobuk & "'"

    isi_semua Rs_Nav
    
    TDB_Rubah.Visible = False
        
        
    Dim n As Object
        For Each n In Me
                        If TypeOf n Is TextBox Then
                        
                         If Not (Left(UCase(n.Name), 10) = UCase("tno_anggota") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan" Or n.Name = "Txt_Kode") Then
                            n.Enabled = True
                         End If
                         
                        End If
            
            If TypeOf n Is DTPicker Then n.Enabled = True
            If TypeOf n Is TDBNumber Then n.Enabled = True
            If TypeOf n Is CheckBox Then n.Enabled = True
            If TypeOf n Is OptionButton Then n.Enabled = True
'            If TypeOf n Is OptionButton Then n.Enabled = True
'            If TypeOf n Is CheckBox Then n.Enabled = True
            
            If TypeOf n Is CommandButton Then
                If n.Caption = "..." Then
                    n.Enabled = False
                End If
            End If
            
        Next

    Cmd_Simpan.Enabled = True
    rubah = True
    
    dtp_tgl.SetFocus
    
End Sub


Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then TDB_Rubah.Visible = False: Cmd_Batal_Click
End Sub

Private Sub t_b_brt_Change()
'    total_harga
End Sub

Private Sub t_b_brt_GotFocus()
    Call Focus_(t_b_brt)
End Sub

Private Sub t_b_brt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_k_brt.SetFocus
End Sub

Private Sub t_b_brt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If

End Sub

Private Sub t_b_brt_LostFocus()
    If t_b_brt.Text = "" Then t_b_brt.Text = 0
End Sub

Private Sub t_b_hrg_GotFocus()
    Call Focus_(t_b_hrg)
End Sub

Private Sub t_b_hrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_k_hrg.SetFocus
End Sub

Private Sub t_b_hrg_KeyPress(KeyAscii As Integer)
'    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
'        KeyAscii = 0
'    End If

End Sub

Private Sub t_b_hrg_LostFocus()
'    If t_b_hrg.Text = "" Then t_b_hrg.SetFocus
End Sub

Private Sub t_b_susut_Change()
    
    Dim tbg_bsh As Double
        If t_b_Timbang.Text = "" Then
            tbg_bsh = 0
        Else
            tbg_bsh = t_b_Timbang.Text
        End If
        
    Dim sst As Double
        If t_b_susut.Text = "" Then
            sst = 0
        Else
            sst = Replace(t_b_susut.Text, ".", ",")
        End If
    
    Dim ssut_p As Double
        ssut_p = tbg_bsh * (sst / 100)
    
        If ssut_p <> 0 Then
            t_bp_susut.Text = FormatNumber(ssut_p, 3)
            
            If t_bp_susut.Text = "" Then
                t_bp_susut.Text = 0
            End If
            
        Else
            t_bp_susut.Text = ssut_p
        End If
    
    total_harga
End Sub

Private Sub t_b_susut_GotFocus()
    Call Focus_(t_b_susut)
End Sub

Private Sub t_b_susut_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_adm.SetFocus
End Sub

Private Sub t_b_susut_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub t_b_susut_LostFocus()
    If t_b_susut.Text = "" Then t_b_susut.Text = 0
End Sub

Private Sub t_b_Timbang_Change()

'    Dim kotor As Double
'        If t_k_timbang.Text = "" Then
'            kotor = 0
'        Else
'            kotor = t_k_timbang
'        End If
'
'    Dim jml As Double
'        jml = kotor * krimawal1
'
'    If kotor = 0 Then
'        tdb_krm1.Value = IIf((krimawal1 = 0), Null, krimawal1)
'    Else
'        tdb_krm1.Value = IIf((jml = 0), Null, jml)
'    End If


    t_b_susut_Change
'    total_harga
End Sub

Private Sub t_b_Timbang_GotFocus()
    Call Focus_(t_b_Timbang)
End Sub

Private Sub t_b_Timbang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ck_wjb.SetFocus
End Sub

Private Sub t_b_Timbang_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub t_b_Timbang_LostFocus()
    If t_b_Timbang.Text = "" Then t_b_Timbang.Text = 0
End Sub

Private Sub t_k_brt_Change()
'    total_harga
End Sub

Private Sub t_k_brt_GotFocus()
    Call Focus_(t_k_brt)
End Sub

Private Sub t_k_brt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_b_hrg.SetFocus
End Sub

Private Sub t_k_brt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If

End Sub

Private Sub t_k_brt_LostFocus()
    If t_k_brt.Text = "" Then t_k_brt.Text = 0
End Sub

Private Sub t_k_hrg_GotFocus()
    Call Focus_(t_k_hrg)
End Sub

Private Sub t_k_hrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_krm.SetFocus
End Sub

Private Sub t_k_hrg_KeyPress(KeyAscii As Integer)
'    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
'        KeyAscii = 0
'    End If

End Sub

Private Sub t_k_hrg_LostFocus()
'    If t_k_hrg.Text = "" Then t_k_hrg.Text = 0
End Sub

Private Sub t_k_susut_Change()

    Dim tbg_krg As Double
        If t_k_timbang.Text = "" Then
            tbg_krg = 0
        Else
            tbg_krg = t_k_timbang.Text
        End If
        
    Dim sst As Double
        If t_k_susut.Text = "" Then
            sst = 0
        Else
            sst = Replace(t_k_susut.Text, ".", ",")
        End If
    
    Dim ssut_k As Double
        ssut_k = tbg_krg * (sst / 100)
        
        If ssut_k <> 0 Then
            t_kp_susut.Text = FormatNumber(ssut_k, 3)
            
            If t_kp_susut.Text = "" Then
                t_kp_susut = 0
            End If
            
        Else
            t_kp_susut.Text = ssut_k
        End If

    total_harga
    
End Sub

Private Sub t_k_susut_GotFocus()
    Call Focus_(t_k_susut)
End Sub

Private Sub t_k_susut_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_b_susut.SetFocus
End Sub

Private Sub t_k_susut_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If

End Sub

Private Sub t_k_susut_LostFocus()
    If t_k_susut.Text = "" Then t_k_susut.Text = 0
End Sub

Private Sub t_k_timbang_Change()
    
'    Dim kotor As Double
'        If t_k_timbang.Text = "" Then
'            kotor = 0
'        Else
'            kotor = t_k_timbang
'        End If
'
'    Dim jml As Double
'        jml = kotor * krimawal
'
'    If kotor = 0 Then
'        tdb_krm.Value = IIf((krimawal = 0), Null, krimawal)
'    Else
'        tdb_krm.Value = IIf((jml = 0), Null, jml)
'    End If
    
    t_k_susut_Change
'    total_harga
    
End Sub

Private Sub t_k_timbang_GotFocus()
    Call Focus_(t_k_timbang)
End Sub

Private Sub t_k_timbang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then t_b_Timbang.SetFocus
End Sub

Private Sub t_k_timbang_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub t_k_timbang_LostFocus()
    If t_k_timbang.Text = "" Then t_k_timbang.Text = 0
End Sub

Private Sub tdb_adm_Change()
'    total_harga
End Sub

Private Sub tdb_adm_GotFocus()
    Call Focus_(tdb_adm)
End Sub

Private Sub tdb_adm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_krm.SetFocus
End Sub

Private Sub TDB_Anggota_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Anggota_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Anggota.Top = TDB_Anggota.Top - (yold - Y)
   TDB_Anggota.Left = TDB_Anggota.Left - (xold - X)
End If

End Sub

Private Sub TDB_Anggota_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub tdb_cc_krd_Change()
    total_harga
End Sub

Private Sub tdb_cc_krd_GotFocus()
    Call Focus_(tdb_cc_krd)
End Sub

Private Sub tdb_cc_krd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdb_dp.SetFocus
    End If
End Sub

Private Sub tdb_cc_ppk_Change()
    total_harga
End Sub

Private Sub tdb_cc_ppk_GotFocus()
    Call Focus_(tdb_cc_ppk)
End Sub

Private Sub tdb_cc_ppk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_cc_krd.SetFocus
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

Private Sub tdb_dp_Change()
    total_harga
End Sub

Private Sub tdb_dp_GotFocus()
    Call Focus_(tdb_dp)
End Sub

Private Sub tdb_dp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub tdb_hrg_br_GotFocus()
    Call Focus_(tdb_hrg_br)
End Sub

Private Sub tdb_hrg_br_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub tdb_krm_Change()
'    total_harga
End Sub

Private Sub tdb_krm_GotFocus()
    Call Focus_(tdb_krm)
End Sub

Private Sub tdb_krm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ck_wjb.SetFocus
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

Private Sub tdb_skr_Change()
    total_harga
End Sub

Private Sub tdb_skr_GotFocus()
    Call Focus_(tdb_skr)
End Sub

Private Sub tdb_skr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_cc_ppk.SetFocus
End Sub

Private Sub tdb_wjb_Change()
    total_harga
End Sub

Private Sub tno_anggota_GotFocus()
    Call Focus_(tno_anggota)
End Sub

Private Sub tno_anggota_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF3 Then cbr_anggota_Click
    If KeyCode = 13 And tno_anggota.Text <> "" Then
        
        
        Dim sql As String
        Dim rs As Recordset
        Dim n As Object
        
            sql = "select no_anggota,nama from tb_anggota where no_anggota='" & Trim(tno_anggota.Text) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon, adOpenKeyset
            
            With rs
                
                If Not .EOF Then
                    
                    tnama_anggota.Text = IIf(Not IsNull(!nama), !nama, "")
                    
                    For Each n In Me
                        
                        If TypeOf n Is TextBox Then
                            If n.Name <> "tno_anggota" Then
                                If Left(n.Name, 3) = "t_b" Or Left(n.Name, 3) = "t_k" Then
                                    n.Enabled = True
                                    n.Text = 0
                                ElseIf Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                                    n.Enabled = True
                                End If
                            End If
                        End If
                        
                        If TypeOf n Is DTPicker Then
                            n.Enabled = True
                            n.Value = Date
                        End If
                        
                        If TypeOf n Is TDBNumber Then
                            n.Enabled = True
                            n.Value = Null
                        End If
                        
                        If TypeOf n Is CheckBox Then
                            n.Enabled = True
                            n.Value = vbUnchecked
                        End If
                        
                        If TypeOf n Is OptionButton Then n.Enabled = False
                        If TypeOf n Is CommandButton Then
                            If n.Caption = "..." Then
                                n.Enabled = True
                            End If
                        End If
                        
                    Next
                    
                    
                    cari_penyusutan
                    cek_harga_perkg
                    cari_biaya_kirim_perkg
                    cari_biaya_adm_perkg
                    
                    Set n = Nothing
                    Cmd_Simpan.Enabled = True
                    dtp_tgl.SetFocus
                    
                Else
                    
                        For Each n In Me
    
                            If TypeOf n Is TextBox Then
                                If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                                    n.Enabled = False
                                End If
                            End If
                            
'                            If TypeOf n Is CheckBox Then n.Enabled = False
                    '        If TypeOf n Is OptionButton Then n.Enabled = False
                            If TypeOf n Is DTPicker Then n.Enabled = False
                            If TypeOf n Is TDBNumber Then n.Enabled = False
                            If TypeOf n Is CheckBox Then n.Enabled = False
                            If TypeOf n Is OptionButton Then n.Enabled = False
                            If TypeOf n Is CommandButton Then
                                If n.Caption = "..." Then
                                    n.Enabled = True
                                End If
                            End If
                                
                        Next
                        
                        Set n = Nothing
                        Cmd_Simpan.Enabled = False
                    
                    
                End If
                
            End With
        
        
    End If
    
End Sub

Private Sub Txt_Cr_Anggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Anggota.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Anggota.Visible = False: tno_anggota.SetFocus
End Sub

Private Sub Txt_Cr_Anggota_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim rs As Recordset
    
        sql = "select top 100 * from tb_anggota"
        
    If Txt_Cr_Anggota(0).Text <> "" Or Txt_Cr_Anggota(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Anggota(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Anggota(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by id desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Set Grid_Anggota.DataSource = rs
        Grid_Anggota.Refresh
    
End Sub

Private Sub Txt_Cr_Daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
        sql = "select top 100 * from Tb_Penimbangan"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" Or Txt_Cr_Daftar(2).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
            Case 2
                If Len(Txt_Cr_Daftar(2).Text) = 10 Then
                    sql = sql & " where tgl_ambil='" & Format(Txt_Cr_Daftar(2).Text, "yyyy/mm/dd") & "'"
                End If
            
        End Select
    End If
    
    sql = sql & " order by id desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Daftar.DataSource = Rs_Nav
        Grid_Daftar.Refresh
    
End Sub

Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
        sql = "select top 100 * from Tb_Penimbangan"
        
    If Txt_Cr_Rubah(0).Text <> "" Or Txt_Cr_Rubah(1).Text <> "" Or Txt_Cr_Rubah(2).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
            Case 2
                If Len(Txt_Cr_Rubah(2).Text) = 10 Then
                    sql = sql & " where tgl_ambil='" & Format(Txt_Cr_Rubah(2).Text, "yyyy/mm/dd") & "'"
                End If
            
        End Select
    End If
    
    sql = sql & " order by id desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Rubah.DataSource = Rs_Nav
        Grid_Rubah.Refresh
    
End Sub
