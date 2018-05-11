VERSION 5.00
Begin VB.Form frm_sel_anggota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi Aggota"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_sel_anggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5220
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
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5175
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
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txt_kab 
            Height          =   320
            Left            =   1215
            TabIndex        =   27
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txt_kec 
            Height          =   320
            Left            =   1200
            TabIndex        =   25
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox tmax2 
            Height          =   320
            Left            =   2415
            TabIndex        =   24
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox tmin2 
            Height          =   320
            Left            =   2400
            TabIndex        =   23
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox tluas2 
            Height          =   320
            Left            =   2415
            TabIndex        =   22
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox tmax1 
            Height          =   320
            Left            =   1200
            TabIndex        =   17
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox tmin1 
            Height          =   320
            Left            =   1185
            TabIndex        =   15
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox Txt_Kode 
            Height          =   320
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox Txt_Nama 
            Height          =   320
            Left            =   1200
            TabIndex        =   6
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox Txt_Alamat 
            Height          =   320
            Left            =   1200
            TabIndex        =   5
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox tluas1 
            Height          =   320
            Left            =   1200
            TabIndex        =   4
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kab :"
            Height          =   195
            Index           =   10
            Left            =   705
            TabIndex        =   28
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kec :"
            Height          =   195
            Index           =   9
            Left            =   705
            TabIndex        =   26
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            Height          =   195
            Index           =   8
            Left            =   2040
            TabIndex        =   21
            Top             =   3000
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            Height          =   195
            Index           =   7
            Left            =   2040
            TabIndex        =   20
            Top             =   2640
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            Height          =   195
            Index           =   6
            Left            =   2040
            TabIndex        =   19
            Top             =   2280
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Max :"
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   18
            Top             =   2880
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil Min :"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Anggota :"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   11
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama :"
            Height          =   195
            Index           =   3
            Left            =   555
            TabIndex        =   10
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desa :"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   9
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Luas Kebun :"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   8
            Top             =   2160
            Width           =   930
         End
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
         TabIndex        =   12
         Top             =   0
         Width           =   3855
         Begin VB.OptionButton Opt_Kriteria 
            Caption         =   "&Berdasarkan Kriteria"
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton Opt_Semua 
            Caption         =   "&Semua"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton Cmd_Lihat 
         Caption         =   "&Tampil"
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   4080
         TabIndex        =   1
         Top             =   3840
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_sel_anggota"
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
    
    sql = "select  * from tb_anggota order by nama asc"
    
    Else
    
    If Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "" Or (tmin1.Text <> "" And tmin2.Text <> "") _
        Or (tluas1.Text <> "" And tluas2.Text <> "") Or (tmax1.Text <> "" And tmax2.Text <> "") Or txt_kec.Text <> "" Or txt_kab.Text <> "" Then
        
        sql = "select * from tb_anggota where"
        
        If Txt_Kode.Text <> "" Then
            sql = sql & " no_anggota like '%" & Trim(Txt_Kode.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text = "" Then
            sql = sql & " nama like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Txt_Nama.Text <> "" And Txt_Kode.Text <> "" Then
            sql = sql & " and nama like '%" & Trim(Txt_Nama.Text) & "%'"
        End If
        
        If Txt_Alamat.Text <> "" And Txt_Nama.Text = "" And Txt_Kode.Text = "" Then
            sql = sql & " desa like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If
        
        If Txt_Alamat.Text <> "" And (Txt_Nama.Text <> "" Or Txt_Kode.Text <> "") Then
            sql = sql & " and desa like '%" & Trim(Txt_Alamat.Text) & "%'"
        End If
            
        If tluas1.Text <> "" And tluas2.Text <> "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" And Txt_Alamat.Text = "" Then
            sql = sql & " Luas_Kebun >=" & Trim(tluas1.Text) & " and Luas_Kebun <=" & Trim(tluas2.Text)
        End If
        
        If tluas1.Text <> "" And tluas2.Text <> "" And (Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "") Then
            sql = sql & " and Luas_Kebun >=" & Trim(tluas1.Text) & " and Luas_Kebun <=" & Trim(tluas2.Text)
        End If
        
        If tmin1.Text <> "" And tmin2.Text <> "" And tluas1.Text = "" And tluas2.Text = "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" And Txt_Alamat.Text = "" Then
            sql = sql & " hasil_min >=" & Trim(tmin1.Text) & " and hasil_min <=" & Trim(tmin2.Text)
        End If
        
        If tmin1.Text <> "" And tmin2.Text <> "" And ((tluas1.Text <> "" And tluas2.Text <> "") Or Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "") Then
            sql = sql & " and hasil_min >=" & Trim(tmin1.Text) & " and hasil_min <=" & Trim(tmin2.Text)
        End If
        
        If tmax1.Text <> "" And tmax2.Text <> "" And tmin1.Text = "" And tmin2.Text = "" And tluas1.Text = "" And tluas2.Text = "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" And Txt_Alamat.Text = "" Then
            sql = sql & " hasil_max >=" & Trim(tmax1.Text) & " and hasil_max <=" & Trim(tmax2.Text)
        End If
        
        If tmax1.Text <> "" And tmax2.Text <> "" And ((tmin1.Text <> "" And tmin2.Text <> "") Or (tluas1.Text <> "" And tluas2.Text <> "") Or Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "") Then
            sql = sql & " and hasil_max >=" & Trim(tmax1.Text) & " and hasil_max <=" & Trim(tmax2.Text)
        End If
        
        If txt_kec.Text <> "" And tmax1.Text = "" And tmax2.Text = "" And tmin1.Text = "" And tmin2.Text = "" And tluas1.Text = "" And tluas2.Text = "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" And Txt_Alamat.Text = "" Then
            sql = sql & " kec like '%" & Trim(txt_kec.Text) & "%'"
        End If
        
        If txt_kec.Text <> "" And ((tmax1.Text <> "" And tmax2.Text <> "") Or (tmin1.Text <> "" And tmin2.Text <> "") Or (tluas1.Text <> "" And tluas2.Text <> "") Or Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "") Then
            sql = sql & " and kec like '%" & Trim(txt_kec.Text) & "%'"
        End If
        
        If txt_kab.Text <> "" And txt_kec.Text = "" And tmax1.Text = "" And tmax2.Text = "" And tmin1.Text = "" And tmin2.Text = "" And tluas1.Text = "" And tluas2.Text = "" And Txt_Kode.Text = "" And Txt_Nama.Text = "" And Txt_Alamat.Text = "" Then
            sql = sql & " kab like '%" & Trim(txt_kab.Text) & "%'"
        End If
        
        If txt_kab.Text <> "" And (txt_kec.Text <> "" Or (tmax1.Text <> "" And tmax2.Text <> "") Or (tmin1.Text <> "" And tmin2.Text <> "") Or (tluas1.Text <> "" And tluas2.Text <> "") Or Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Txt_Alamat.Text <> "") Then
            sql = sql & " kab and like '%" & Trim(txt_kab.Text) & "%'"
        End If
            
        sql = sql & " order by nama asc"
        
        
    Else
        
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Kriteria pencarian harus diisi", vbOKOnly + vbInformation, "Informasi"))
        
        Exit Sub
    End If
    
    End If
    
'    khusus_user = Mid(Utama.StatusBar1.Panels(5).Text, 7, Len(Utama.StatusBar1.Panels(5).Text))
    
    Mysq = sql
    
    Load frm_lap_anggota_Detail
        frm_lap_anggota_Detail.Show
    
    
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
    If KeyCode = 13 Then tmin1.SetFocus
End Sub

Private Sub tmax1_GotFocus()
    Call Focus_(tmax1)
End Sub

Private Sub tmax1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tmax2.SetFocus
End Sub

Private Sub tmax2_GotFocus()
    Call Focus_(tmax2)
End Sub

Private Sub tmax2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub

Private Sub tmin1_GotFocus()
    Call Focus_(tmin1)
End Sub

Private Sub tmin1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tmin2.SetFocus
End Sub

Private Sub tmin2_GotFocus()
    Call Focus_(tmin2)
End Sub

Private Sub tmin2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tmax1.SetFocus
End Sub

Private Sub Txt_Alamat_GotFocus()
    Call Focus_(Txt_Alamat)
End Sub

Private Sub Txt_Alamat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_kec.SetFocus
End Sub


Private Sub txt_kab_GotFocus()
    Call Focus_(txt_kab)
End Sub

Private Sub txt_kab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tluas1.SetFocus
End Sub

Private Sub txt_kec_GotFocus()
    Call Focus_(txt_kec)
End Sub

Private Sub txt_kec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_kab.SetFocus
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
    If KeyCode = 13 Then Txt_Alamat.SetFocus
End Sub
