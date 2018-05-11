VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_sel_penimbangan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Penimbangan"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_sel_penimbangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5295
         Begin VB.TextBox tluas1 
            Height          =   320
            Left            =   1560
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Txt_Nama 
            Height          =   320
            Left            =   1560
            TabIndex        =   4
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox Txt_Kode 
            Height          =   320
            Left            =   1560
            TabIndex        =   3
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox tluas2 
            Height          =   320
            Left            =   2775
            TabIndex        =   2
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSMask.MaskEdBox txt_tgl1 
            Height          =   315
            Left            =   1560
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_tgl2 
            Height          =   315
            Left            =   3360
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            Height          =   195
            Index           =   18
            Left            =   3000
            TabIndex        =   18
            Top             =   1080
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Bersih :"
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   10
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Pengambilan :"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   9
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama :"
            Height          =   195
            Index           =   3
            Left            =   915
            TabIndex        =   8
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
            Left            =   465
            TabIndex        =   7
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            Height          =   195
            Index           =   6
            Left            =   2400
            TabIndex        =   6
            Top             =   1500
            Visible         =   0   'False
            Width           =   225
         End
      End
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   615
         Left            =   4440
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Lihat 
         Caption         =   "&Tampil"
         Height          =   615
         Left            =   3480
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   3855
         Begin VB.OptionButton Opt_Kriteria 
            Caption         =   "&Berdasarkan Kriteria"
            Height          =   255
            Left            =   960
            TabIndex        =   12
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton Opt_Semua 
            Caption         =   "&Semua"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frm_sel_penimbangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Foto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub

Private Sub Cmd_Keluar_Click()

    Unload Me

End Sub

Private Sub Cmd_Lihat_Click()
    
    Dim sql As String
    
    If Opt_Semua.Value = True Then
    
    sql = "select  * from VIEW_Penimbangan order by tgl_ambil desc"
    
    Else
    
    If Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or _
        (tluas1.Text <> "" And tluas2.Text <> "") Or (txt_tgl1.Text <> "__/__/____" And txt_tgl2.Text <> "__/__/____") Then
        
        sql = "select * from VIEW_Penimbangan where"
        
        If Txt_Kode.Text <> "" Then
            sql = sql & " no_anggota like '%" & Trim(Txt_Kode.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text = "" Then
            sql = sql & " nama like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text <> "" Then
            sql = sql & " and nama like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
'        If tluas1.Text <> "" And tluas2.Text <> "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" Then
'            sql = sql & " Luas_Kebun >=" & Trim(tluas1.Text) & " and Luas_Kebun <=" & Trim(tluas2.Text)
'        End If
'
'        If tluas1.Text <> "" And tluas2.Text <> "" And (Txt_Kode.Text <> "" Or Txt_Nama.Text <> "") Then
'            sql = sql & " and Luas_Kebun >=" & Trim(tluas1.Text) & " and Luas_Kebun <=" & Trim(tluas2.Text)
'        End If
        
        If txt_tgl1.Text <> "__/__/____" And txt_tgl2.Text <> "__/__/____" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " tgl_ambil >='" & Format(txt_tgl1.Text, "yyyy/mm/dd") & "' and tgl_ambil <='" & Format(txt_tgl2.Text, "yyyy/mm/dd") & "'"
        End If
        
        If txt_tgl1.Text <> "__/__/____" And txt_tgl2.Text <> "__/__/____" And (Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and tgl_ambil >='" & Format(txt_tgl1.Text, "yyyy/mm/dd") & "' and tgl_ambil <='" & Format(txt_tgl2.Text, "yyyy/mm/dd") & "'"
        End If
        
        sql = sql & " order by tgl_ambil desc"
        
        
    Else
        
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Kriteria pencarian harus diisi", vbOKOnly + vbInformation, "Informasi"))
        
        Exit Sub
    End If
    
    End If
    
'    khusus_user = Mid(Utama.StatusBar1.Panels(5).Text, 7, Len(Utama.StatusBar1.Panels(5).Text))
    
    Mysq = sql
    
    Load frm_lap_penimbangan
        frm_lap_penimbangan.Show
    
    
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
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 250
    End With
    
    Opt_Semua.Value = True

'' akses command ''

'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Lihat.Enabled = c_laporan

'' stop here ''


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If
    
 
End Sub

Private Sub Opt_Kriteria_Click()
    
    If Opt_Kriteria.Value = True Then Frame2.Enabled = True
    
End Sub

Private Sub Opt_Kriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Kode.SetFocus
End Sub

Private Sub Opt_Semua_Click()
    If Opt_Semua.Value = True Then
        Frame2.Enabled = False
    
    Dim a As Object
        For Each a In Me
            If TypeOf a Is TextBox Then
                a.Text = ""
            End If
            
            If TypeOf a Is MaskEdBox Then a.Text = "__/__/____"
        Next
        
        Set a = Nothing
    
    End If
End Sub

Private Sub Opt_Semua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Lihat.Enabled = True Then Cmd_Lihat.SetFocus
    End If
End Sub

Private Sub Tgl_Masuk2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Cmd_Lihat.Enabled = True Then Cmd_Lihat.SetFocus
    End If
        
End Sub

Private Sub tluas1_GotFocus()
    Call Focus_(tluas1)
End Sub

Private Sub tluas1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tluas2.SetFocus
End Sub

Private Sub tluas2_GotFocus()
    Call Focus_(tluas2)
End Sub

Private Sub tluas2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub





Private Sub txt_kode_GotFocus()
    Call Focus_(Txt_Kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Nama.SetFocus
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_tgl1.SetFocus
End Sub


Private Sub txt_tgl1_GotFocus()
    Call Focus_(txt_tgl1)
End Sub

Private Sub txt_tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_tgl2.SetFocus
End Sub

Private Sub txt_tgl2_GotFocus()
    Call Focus_(txt_tgl2)
End Sub

Private Sub txt_tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub
