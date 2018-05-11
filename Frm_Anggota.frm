VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form Frm_Anggota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Anggota"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Anggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3855
      Left            =   2160
      TabIndex        =   68
      Top             =   2880
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Anggota.frx":08CA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Anggota.frx":08E6
      Childs          =   "Frm_Anggota.frx":0992
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   71
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   70
         Top             =   600
         Width           =   2535
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
         TabIndex        =   69
         Top             =   360
         Width           =   6015
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2775
         Left            =   240
         OleObjectBlob   =   "Frm_Anggota.frx":09AE
         TabIndex        =   72
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   40
         Left            =   360
         TabIndex        =   75
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   41
         Left            =   3120
         TabIndex        =   74
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   73
         Top             =   120
         Width           =   870
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Hapus 
      Height          =   3615
      Left            =   -4800
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6376
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Anggota.frx":3932
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Anggota.frx":394E
      Childs          =   "Frm_Anggota.frx":39FA
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   63
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   62
         Top             =   600
         Width           =   2655
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
         TabIndex        =   61
         Top             =   360
         Width           =   5895
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Hapus 
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "Frm_Anggota.frx":3A16
         TabIndex        =   64
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   38
         Left            =   2880
         TabIndex        =   67
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   39
         Left            =   240
         TabIndex        =   66
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   120
         Width           =   870
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   3735
      Left            =   -4320
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6588
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Anggota.frx":6999
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Anggota.frx":69B5
      Childs          =   "Frm_Anggota.frx":6A61
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   55
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   54
         Top             =   600
         Width           =   2655
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
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   5775
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Frm_Anggota.frx":6A7D
         TabIndex        =   56
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   37
         Left            =   2880
         TabIndex        =   58
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   120
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      FillColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8745
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   -240
      Width           =   8895
      Begin VB.Frame Frame1 
         Caption         =   "Alamat"
         Height          =   1815
         Left            =   1560
         TabIndex        =   84
         Top             =   1800
         Width           =   4215
         Begin VB.TextBox tkab 
            Height          =   300
            Left            =   840
            TabIndex        =   91
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox tkec 
            Height          =   300
            Left            =   840
            TabIndex        =   89
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox tdesa 
            Height          =   300
            Left            =   840
            TabIndex        =   87
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox talamat 
            Height          =   300
            Left            =   840
            TabIndex        =   85
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kab :"
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   92
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kec :"
            Height          =   195
            Index           =   28
            Left            =   360
            TabIndex        =   90
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desa :"
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   88
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jl :"
            Height          =   195
            Index           =   26
            Left            =   510
            TabIndex        =   86
            Top             =   360
            Width           =   210
         End
      End
      Begin VB.TextBox tlahir 
         Height          =   300
         Left            =   2400
         TabIndex        =   80
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "&Cetak Kartu Anggota"
         Height          =   735
         Left            =   7920
         TabIndex        =   77
         Top             =   7905
         Width           =   855
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
         Left            =   3360
         TabIndex        =   44
         Top             =   7800
         Width           =   4455
         Begin VB.CommandButton Cmd_Keluar 
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   3480
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Daftar 
            Caption         =   "&Daftar"
            Height          =   495
            Left            =   2640
            TabIndex        =   46
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Hapus 
            Caption         =   "&Hapus"
            Height          =   495
            Left            =   1800
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Rubah 
            Caption         =   "&Rubah"
            Height          =   495
            Left            =   960
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Tambah 
            Caption         =   "&Tambah"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Simpan 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Cmd_Batal 
            Caption         =   "&Batal"
            Height          =   495
            Left            =   960
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
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
         TabIndex        =   39
         Top             =   7800
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
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox tpend 
         Height          =   300
         Left            =   2400
         TabIndex        =   27
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox tistri 
         Height          =   300
         Left            =   2400
         TabIndex        =   25
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox tktp 
         Height          =   300
         Left            =   2400
         TabIndex        =   24
         Top             =   4440
         Width           =   3015
      End
      Begin VB.ComboBox cbo_status 
         Height          =   315
         ItemData        =   "Frm_Anggota.frx":9A00
         Left            =   2400
         List            =   "Frm_Anggota.frx":9A0A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox tkerjaan 
         Height          =   300
         Left            =   2400
         TabIndex        =   22
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox tnama 
         Height          =   300
         Left            =   2400
         TabIndex        =   21
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox tno_anggota4 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox tno_anggota3 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox tno_anggota2 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox tno_anggota1 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin TDBNumber6Ctl.TDBNumber tdbanak 
         Height          =   300
         Left            =   2400
         TabIndex        =   26
         Top             =   5160
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9A26
         Caption         =   "Frm_Anggota.frx":9A46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":9AB2
         Keys            =   "Frm_Anggota.frx":9AD0
         Spin            =   "Frm_Anggota.frx":9B1A
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
      Begin TDBNumber6Ctl.TDBNumber tdbluas 
         Height          =   300
         Left            =   2400
         TabIndex        =   28
         Top             =   5880
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9B42
         Caption         =   "Frm_Anggota.frx":9B62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":9BCE
         Keys            =   "Frm_Anggota.frx":9BEC
         Spin            =   "Frm_Anggota.frx":9C36
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
      Begin TDBNumber6Ctl.TDBNumber tdbproduksi 
         Height          =   300
         Left            =   2400
         TabIndex        =   31
         Top             =   6240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9C5E
         Caption         =   "Frm_Anggota.frx":9C7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":9CEA
         Keys            =   "Frm_Anggota.frx":9D08
         Spin            =   "Frm_Anggota.frx":9D52
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
      Begin TDBNumber6Ctl.TDBNumber tdbhasilmin 
         Height          =   300
         Left            =   2400
         TabIndex        =   34
         Top             =   6600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9D7A
         Caption         =   "Frm_Anggota.frx":9D9A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":9E06
         Keys            =   "Frm_Anggota.frx":9E24
         Spin            =   "Frm_Anggota.frx":9E6E
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
      Begin TDBNumber6Ctl.TDBNumber tdbhasilmax 
         Height          =   300
         Left            =   2400
         TabIndex        =   35
         Top             =   6960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9E96
         Caption         =   "Frm_Anggota.frx":9EB6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":9F22
         Keys            =   "Frm_Anggota.frx":9F40
         Spin            =   "Frm_Anggota.frx":9F8A
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
      Begin TDBNumber6Ctl.TDBNumber tdbbruh 
         Height          =   300
         Left            =   2400
         TabIndex        =   38
         Top             =   7320
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "Frm_Anggota.frx":9FB2
         Caption         =   "Frm_Anggota.frx":9FD2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_Anggota.frx":A03E
         Keys            =   "Frm_Anggota.frx":A05C
         Spin            =   "Frm_Anggota.frx":A0A6
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
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   300
         Left            =   2400
         TabIndex        =   81
         Top             =   1440
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
         Format          =   51118081
         CurrentDate     =   39372
      End
      Begin VB.Label LUmur 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5160
         TabIndex        =   83
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Umur :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   25
         Left            =   4440
         TabIndex        =   82
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lhr :"
         Height          =   195
         Index           =   24
         Left            =   1695
         TabIndex        =   79
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lhr :"
         Height          =   195
         Index           =   23
         Left            =   1365
         TabIndex        =   78
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8160
         TabIndex        =   76
         Top             =   7560
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   22
         Left            =   3240
         TabIndex        =   37
         Top             =   7080
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         Height          =   195
         Index           =   21
         Left            =   3240
         TabIndex        =   36
         Top             =   6720
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "SansSerif"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   20
         Left            =   3375
         TabIndex        =   33
         Top             =   6240
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         Height          =   195
         Index           =   19
         Left            =   3240
         TabIndex        =   32
         Top             =   6315
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "SansSerif"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   18
         Left            =   3375
         TabIndex        =   30
         Top             =   5925
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         Height          =   195
         Index           =   17
         Left            =   3240
         TabIndex        =   29
         Top             =   6000
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   16
         Left            =   4440
         TabIndex        =   19
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   15
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   14
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Karyawan/Buruh :"
         Height          =   195
         Index           =   13
         Left            =   780
         TabIndex        =   13
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Max Perbulan :"
         Height          =   195
         Index           =   12
         Left            =   900
         TabIndex        =   12
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Min Perbulan :"
         Height          =   195
         Index           =   11
         Left            =   945
         TabIndex        =   11
         Top             =   6600
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luas Yang Sudah Produksi :"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   10
         Top             =   6240
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luas Tot Kebun Karet :"
         Height          =   195
         Index           =   9
         Left            =   690
         TabIndex        =   9
         Top             =   5880
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pendidikan Terakhir :"
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   8
         Top             =   5520
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Anak :"
         Height          =   195
         Index           =   7
         Left            =   1605
         TabIndex        =   7
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Istri :"
         Height          =   195
         Index           =   6
         Left            =   1500
         TabIndex        =   6
         Top             =   4800
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP :"
         Height          =   195
         Index           =   5
         Left            =   1665
         TabIndex        =   5
         Top             =   4440
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Index           =   4
         Left            =   1770
         TabIndex        =   4
         Top             =   4080
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan :"
         Height          =   195
         Index           =   3
         Left            =   1515
         TabIndex        =   3
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   195
         Index           =   1
         Left            =   1830
         TabIndex        =   2
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Anggota :"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Frm_Anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rubah As Boolean
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub isi_semua(ByVal rec As Recordset)
    
    With rec
        
        If .EOF Then .MoveLast
        If .BOF Then .MoveFirst
        
        Dim no_anggota1, no_anggota2, no_anggota3, no_anggota4 As String
        Dim no_anggota As String
            
            no_anggota = IIf(Not IsNull(!no_anggota), !no_anggota, "")
            
            If no_anggota <> "" Then
                
                no_anggota1 = Left(no_anggota, 1)
                no_anggota2 = Mid(no_anggota, 3, 2)
                no_anggota3 = Mid(no_anggota, 6, 2)
                no_anggota4 = Right(no_anggota, 5)
                
            Else
                
                no_anggota1 = ""
                no_anggota2 = ""
                no_anggota3 = ""
                no_anggota4 = ""
                
            End If
            
            tno_anggota1.Text = no_anggota1
            tno_anggota2.Text = no_anggota2
            tno_anggota3.Text = no_anggota3
            tno_anggota4.Text = no_anggota4
            
            tnama.Text = IIf(Not IsNull(!NAMA), !NAMA, "")
            tlahir.Text = IIf(Not IsNull(!tempat_lhr), !tempat_lhr, "")
            dtp_tgl.Value = IIf(Not IsNull(!tgl_lhr), !tgl_lhr, Date)
            
            dtp_tgl_Change
            
            talamat.Text = IIf(Not IsNull(!jl), !jl, "")
            tdesa.Text = IIf(Not IsNull(!desa), !desa, "")
            tkec.Text = IIf(Not IsNull(!kec), !kec, "")
            tkab.Text = IIf(Not IsNull(!kab), !kab, "")
            tkerjaan.Text = IIf(Not IsNull(!Pekerjaan), !Pekerjaan, "")
            
            Dim stat As String
                stat = IIf(Not IsNull(!status), !status, "")
            
            If UCase(stat) = UCase("Menikah") Then
                cbo_status.ListIndex = 0
            Else
                cbo_status.ListIndex = 1
            End If
            
            tktp.Text = IIf(Not IsNull(!No_Ktp), !No_Ktp, "")
            tistri.Text = IIf(Not IsNull(!Nama_Istri), !Nama_Istri, "")
            tdbanak.Value = IIf(Not IsNull(!Jml_Anak), !Jml_Anak, Null)
            tpend.Text = IIf(Not IsNull(!Pendidikan_Terakhir), !Pendidikan_Terakhir, "")
            tdbluas.Value = IIf(Not IsNull(!Luas_Kebun), !Luas_Kebun, Null)
            tdbproduksi.Value = IIf(Not IsNull(!Luas_Prod), !Luas_Prod, Null)
            tdbhasilmin.Value = IIf(Not IsNull(!Hasil_Min), !Hasil_Min, Null)
            tdbhasilmax.Value = IIf(Not IsNull(!Hasil_Max), !Hasil_Max, Null)
            tdbbruh.Value = IIf(Not IsNull(!Jml_Kary), !Jml_Kary, Null)
                
            
            If .RecordCount = 0 Then
                Lbl_Info.Caption = "Record Ke " & 0 & " Dari " & .RecordCount & " Record"
            Else
                Lbl_Info.Caption = "Record Ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
            End If
            
    End With
    
End Sub

Private Sub isi_no_anggota()
    
    Dim sql As String
    Dim rs As Recordset
    Dim a As Long
        a = 1
    
        sql = "select no_urut from Tb_Anggota order by no_urut desc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    Dim no_akhir As Double
        no_akhir = 0
    With rs
        
        Do While Not .EOF
            
            Dim no_k As String
                no_k = Right(!no_urut, 5)
            
            Dim no_d As Double
                no_d = CDbl(no_k)
                
                If no_d > no_akhir Then
                    no_akhir = no_d
                End If
        
        a = a + 1
        
        If a >= 100 Then Exit Do
        
        .MoveNext
        Loop
        
    End With
    
    If no_akhir = 0 Then
        tno_anggota4.Text = "00001"
    Else
        
        no_akhir = no_akhir + 1
        
        Dim hsl_n As String
            hsl_n = no_akhir
            
        If Len(hsl_n) = 1 Then
            hsl_n = "0000" & hsl_n
        ElseIf Len(hsl_n) = 2 Then
            hsl_n = "000" & hsl_n
        ElseIf Len(hsl_n) = 3 Then
            hsl_n = "00" & hsl_n
        ElseIf Len(hsl_n) = 4 Then
            hsl_n = "0" & hsl_n
        Else
            hsl_n = hsl_n
        End If
        
        tno_anggota4.Text = hsl_n
            
    End If
    
End Sub

Private Sub cbo_status_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tktp.SetFocus
End Sub

Private Sub Cmd_Batal_Click()

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
        
        cmd_cetak.Enabled = True
        
        Dim n As Object
            
For Each n In Me

        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is DTPicker Then n.Enabled = False
        If TypeOf n Is ComboBox Then n.Enabled = False
        
        If TypeOf n Is TDBNumber Then
            n.Enabled = False
        End If
        
        
        If TypeOf n Is TDBContainer3D Then n.Visible = False
        If TypeOf n Is OptionButton Then n.Enabled = False

Next

Set n = Nothing

LUmur.Caption = "0 Thn"
Txt_Cr_Daftar_KeyUp 0, 0, 0

Cmd_Navigasi_Click 0

    Cmd_Tambah.SetFocus
    
End Sub

Private Sub cmd_cetak_Click()
    
    If tno_anggota1.Text = "" Or tno_anggota2.Text = "" Or tno_anggota3.Text = "" _
        Or tno_anggota4.Text = "" Then Exit Sub
    
    Dim no_anggota As String
    no_anggota = Trim(tno_anggota1.Text) & "/" & Trim(tno_anggota2.Text) & "/" & Trim(tno_anggota3.Text) & "/" & Trim(tno_anggota4.Text)

    
    Mysq = "select * from tb_anggota where no_anggota='" & no_anggota & "'"
    
    frm_lap_kartu.Show
    
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
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_cetak.Enabled = False
    
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
   
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
 
   
    Cmd_Tambah.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_cetak.Enabled = False
    
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
    
    .Left = Me.Width / 2 - .Width / 2
    .Top = Me.Height / 2 - .Height / 2
    
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    Cmd_Rubah.Visible = False
    Cmd_Batal.Visible = True
    Cmd_Hapus.Enabled = False
    Cmd_Daftar.Enabled = False
    Cmd_Keluar.Enabled = False
    
    cmd_cetak.Enabled = False
    
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

kon.BeginTrans
Dim sql, sql1 As String
Dim rs As Recordset
Dim rs1 As Recordset

Dim konfirm As Integer

Dim no_anggota As String
    no_anggota = Trim(tno_anggota1.Text) & "/" & Trim(tno_anggota2.Text) & "/" & Trim(tno_anggota3.Text) & "/" & Trim(tno_anggota4.Text)

Dim luas As Double
    If tdbluas.ValueIsNull Then
        luas = 0
    Else
        luas = Replace(Trim(tdbluas.Value), ",", "")
    End If

Dim prod As Double
    If tdbproduksi.ValueIsNull Then
        prod = 0
    Else
        prod = Replace(Trim(tdbproduksi.Value), ",", "")
    End If

Dim hasilmin As Double
    If tdbhasilmin.ValueIsNull Then
        hasilmin = 0
    Else
        hasilmin = Replace(Trim(tdbhasilmin.Value), ",", "")
    End If

Dim hasilmax As Double
    If tdbhasilmax.ValueIsNull Then
        hasilmax = 0
    Else
        hasilmax = Replace(Trim(tdbhasilmax.Value), ",", "")
    End If

Dim jmlbruh As Double
    If tdbbruh.ValueIsNull Then
        jmlbruh = 0
    Else
        jmlbruh = Replace(Trim(tdbbruh.Value), ",", "")
    End If

Dim jmlank As Double
    If tdbanak.ValueIsNull Then
        jmlank = 0
    Else
        jmlank = Replace(Trim(tdbanak.Value), ",", "")
    End If

If rubah = False Then
    
    
        
        sql = "insert into Tb_Anggota (No_Anggota,Nama,jl,Pekerjaan,Status,No_Ktp,Nama_Istri,Jml_Anak,Pendidikan_Terakhir,Luas_Kebun,Luas_Prod,Hasil_Min,Hasil_Max,Jml_Kary,Tempat_Lhr,Tgl_Lhr,jml_wajib,jml_sukarela,desa,kec,kab,no_urut)"
        sql = sql & " values('" & no_anggota & "','" & Trim(tnama.Text) & "','" & Trim(talamat.Text) & "','" & Trim(tkerjaan.Text) & "','" & cbo_status.Text & "','" & Trim(tktp.Text) & "','" & Trim(tistri.Text) & "'"
        sql = sql & "," & jmlank & ",'" & Trim(tpend.Text) & "'," & luas & "," & prod & "," & hasilmin & "," & hasilmax & "," & jmlbruh & ",'" & Trim(tlahir.Text) & "','" & Format(dtp_tgl.Value, "yyyy/mm/dd") & "',0,0,'" & Trim(tdesa.Text) & "','" & Trim(tkec.Text) & "','" & Trim(tkab.Text) & "'," & Right(no_anggota, 5) & " )"
        
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        kon.CommitTrans
        
        konfirm = CInt(MsgBox("Data karyawan telah disimpan ...", vbOKOnly + vbInformation, "Informasi"))
        
        Cmd_Batal_Click
        
    
Else

    sql = "update Tb_Anggota set nama='" & Trim(tnama.Text) & "',jl='" & Trim(talamat.Text) & "',pekerjaan='" & Trim(tkerjaan.Text) & "',status='" & cbo_status.Text & "',no_ktp='" & Trim(tktp.Text) & "'"
    sql = sql & ",nama_istri='" & Trim(tistri.Text) & "',jml_anak=" & jmlank & ",pendidikan_terakhir='" & Trim(tpend.Text) & "',luas_kebun=" & luas & ",luas_prod=" & prod & ",hasil_min=" & hasilmin & ",hasil_max=" & hasilmax & ",jml_kary=" & jmlbruh
    sql = sql & ",Tempat_Lhr='" & Trim(tlahir.Text) & "',Tgl_Lhr='" & Format(dtp_tgl.Value, "yyyy/mm/dd") & "',desa='" & Trim(tdesa.Text) & "',kec='" & Trim(tkec.Text) & "',kab='" & Trim(tkab.Text) & "' where no_anggota='" & no_anggota & "'"
        
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
    
     cmd_cetak.Enabled = False
    
     tno_anggota1.Enabled = True
     tno_anggota2.Enabled = True
     tno_anggota3.Enabled = True
     tno_anggota4.Enabled = True
        
     tno_anggota1.Text = ""
     tno_anggota2.Text = ""
     tno_anggota3.Text = ""
     tno_anggota4.Text = ""
        
     isi_no_anggota
        
     tno_anggota1.SetFocus
        
End Sub

Private Sub dtp_tgl_Change()
On Error Resume Next
    
    Dim thnnow As Double
    Dim thnlhr As Double
        
        thnnow = Right(Format(Date, "dd/mm/yyyy"), 4)
        thnlhr = Right(Format(dtp_tgl.Value, "dd/mm/yyyy"), 4)
        
    Dim umurnow As Double
        umurnow = thnnow - thnlhr
    
    LUmur.Caption = umurnow & " Thn"
    
End Sub

Private Sub dtp_tgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then talamat.SetFocus
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
        If TypeOf n Is ComboBox Then n.Enabled = False
        If TypeOf n Is OptionButton Then n.Enabled = False
        If TypeOf n Is CommandButton Then
            If n.Caption = "..." Then
                n.Enabled = False
            End If
        End If
            
    Next

Set n = Nothing

LUmur.Caption = "0 Thn"
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

Private Sub grid_daftar_DblClick()
    
    If Grid_Daftar.Row < 0 Then Exit Sub
    
    Dim anggot As String
        anggot = Grid_Daftar.Columns(0).Text
    
    Rs_Nav.MoveFirst
    
    Rs_Nav.Find "no_anggota='" & anggot & "'"

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


Private Sub Grid_Hapus_DblClick()
    
On Error GoTo err_handler
    
    If Grid_Hapus.Row < 0 Then Exit Sub
    
    kon.BeginTrans
    
    If MsgBox("Yakin akan hapus : " & Grid_Hapus.Columns(0).Text & " ...?", vbYesNo + vbQuestion, "Hapus") = vbNo Then
        kon.RollbackTrans
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "delete from tb_anggota where no_anggota='" & Grid_Hapus.Columns(0).Text & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
        
        kon.CommitTrans
        Dim konfirm As Integer
            
            konfirm = CInt(MsgBox(Grid_Hapus.Columns(0).Text & " Telah dihapus", vbOKOnly + vbInformation, "Hapus"))
            
            Txt_Cr_Hapus_KeyUp 0, 0, 0
        
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

If Grid_Rubah.Row < 0 Then Exit Sub

    Txt_Cr_Rubah(0).Text = Grid_Rubah.Columns(0).Text

    Txt_Cr_Rubah_KeyUp 0, 0, 0

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
    
    tnama.SetFocus
    
End Sub


Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then TDB_Rubah.Visible = False: Cmd_Batal_Click
End Sub


Private Sub talamat_GotFocus()
    Call Focus_(talamat)
End Sub

Private Sub talamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdesa.SetFocus
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

Private Sub tdbanak_GotFocus()
    Call Focus_(tdbanak)
End Sub

Private Sub tdbanak_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tpend.SetFocus
End Sub

Private Sub tdbbruh_GotFocus()
    Call Focus_(tdbbruh)
End Sub

Private Sub tdbbruh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Simpan.Enabled = True Then Cmd_Simpan.SetFocus
    End If
End Sub

Private Sub tdbhasilmax_GotFocus()
    Call Focus_(tdbhasilmax)
End Sub

Private Sub tdbhasilmax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbbruh.SetFocus
End Sub

Private Sub tdbhasilmin_GotFocus()
    Call Focus_(tdbhasilmin)
End Sub

Private Sub tdbhasilmin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbhasilmax.SetFocus
End Sub

Private Sub tdbluas_GotFocus()
    Call Focus_(tdbluas)
End Sub

Private Sub tdbluas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbproduksi.SetFocus
End Sub

Private Sub tdbproduksi_GotFocus()
    Call Focus_(tdbproduksi)
End Sub

Private Sub tdbproduksi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbhasilmin.SetFocus
End Sub

Private Sub tdesa_GotFocus()
    Call Focus_(tdesa)
End Sub

Private Sub tdesa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tkec.SetFocus
End Sub

Private Sub tistri_GotFocus()
    Call Focus_(tistri)
End Sub

Private Sub tistri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbanak.SetFocus
End Sub

Private Sub tkab_GotFocus()
    Call Focus_(tkab)
End Sub

Private Sub tkab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tkerjaan.SetFocus
End Sub

Private Sub tkec_GotFocus()
    Call Focus_(tkec)
End Sub

Private Sub tkec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tkab.SetFocus
End Sub

Private Sub tkerjaan_GotFocus()
    Call Focus_(tkerjaan)
End Sub

Private Sub tkerjaan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cbo_status.SetFocus
End Sub

Private Sub tktp_GotFocus()
    Call Focus_(tktp)
End Sub

Private Sub tktp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tistri.SetFocus
End Sub

Private Sub tlahir_GotFocus()
    Call Focus_(tlahir)
End Sub

Private Sub tlahir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtp_tgl.SetFocus
End Sub

Private Sub tnama_GotFocus()
    Call Focus_(tnama)
End Sub

Private Sub tnama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tlahir.SetFocus
End Sub

Private Sub tno_anggota1_GotFocus()
    Call Focus_(tno_anggota1)
End Sub

Private Sub tno_anggota1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tno_anggota2.SetFocus
End Sub

Private Sub tno_anggota1_LostFocus()
    If tno_anggota1.Text <> "" Then tno_anggota1.Text = UCase(tno_anggota1.Text)
End Sub

Private Sub tno_anggota2_GotFocus()
    Call Focus_(tno_anggota2)
End Sub

Private Sub tno_anggota2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tno_anggota3.SetFocus
End Sub

Private Sub tno_anggota3_GotFocus()
    Call Focus_(tno_anggota3)
End Sub

Private Sub tno_anggota3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        
        Dim sql As String
        Dim rs As Recordset
        Dim n As Object
        
        Dim no_anggota As String
            no_anggota = Trim(tno_anggota1.Text) & "/" & Trim(tno_anggota2.Text) & "/" & Trim(tno_anggota3.Text) & "/" & Trim(tno_anggota4.Text)
        
            sql = "select no_anggota from tb_anggota where no_anggota='" & no_anggota & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon, adOpenKeyset
            
     If Not rs.EOF Then
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Kode Sudah ada ...", vbOKOnly + vbInformation, "Informasi"))

    For Each n In Me
    
        If TypeOf n Is TextBox Then
            If Left(UCase(n.Name), 6) <> UCase("Txt_Cr") And Left(UCase(n.Name), 11) <> UCase("tno_anggota") Then
                n.Enabled = False
            End If
        End If
        
        If TypeOf n Is DTPicker Then n.Enabled = False
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
                
                LUmur.Caption = "0 Thn"
                tno_anggota3.SetFocus
                Cmd_Simpan.Enabled = False
                 
                On Error GoTo 0
                Exit Sub
        Else
                    
                    For Each n In Me
                    
                        If TypeOf n Is TextBox Then
                        
                         If Not (UCase(n.Name) = UCase("Txt_Kode") Or n.Name = "Txt_Agama" Or n.Name = "Txt_Status" Or n.Name = "Txt_Pendidikan" Or n.Name = "Txt_Jabatan") Then
                            n.Enabled = True
                         End If
                         
                         If Not (Left(UCase(n.Name), 11) = UCase("tno_anggota")) Then
                            n.Text = ""
                         End If
                         
                        End If
                        
                       If TypeOf n Is DTPicker Then n.Enabled = True
                        If TypeOf n Is TDBNumber Then
                            n.Enabled = True
                            n.Value = Null
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
                    
                LUmur.Caption = "0 Thn"
                Cmd_Simpan.Enabled = True
                tnama.SetFocus
                
        End If
            
    End If

End Sub

Private Sub tpend_GotFocus()
    Call Focus_(tpend)
End Sub

Private Sub tpend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdbluas.SetFocus
End Sub

Private Sub Txt_Cr_Daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Daftar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
           
    Dim sql As String
        sql = "select top 100 * from tb_anggota"
        
    If Txt_Cr_Daftar(0).Text <> "" Or Txt_Cr_Daftar(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Daftar(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Daftar(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by id desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Daftar.DataSource = Rs_Nav
        Grid_Daftar.Refresh

End Sub

Private Sub Txt_Cr_Hapus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Hapus_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sql As String
        sql = "select top 100 * from tb_anggota"
            
    If Txt_Cr_Hapus(0).Text <> "" Or Txt_Cr_Hapus(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Hapus(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Hapus(1).Text) & "%'"
        End Select
    End If

    sql = sql & " order by id desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Hapus.DataSource = Rs_Nav
        Grid_Hapus.Refresh
    
End Sub


Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then Cmd_Batal_Click
End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)


            
    Dim sql As String
        sql = "select top 100 * from tb_anggota"
        
    If Txt_Cr_Rubah(0).Text <> "" Or Txt_Cr_Rubah(1).Text <> "" Then
        Select Case Index
            Case 0
                sql = sql & " where no_anggota like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
            Case 1
                sql = sql & " where nama like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
        End Select
    End If
    
    sql = sql & " order by id desc"
    
    Set Rs_Nav = New ADODB.Recordset
        Rs_Nav.Open sql, kon, adOpenKeyset
    
    Set Grid_Rubah.DataSource = Rs_Nav
        Grid_Rubah.Refresh
    
End Sub
