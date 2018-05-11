VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Utama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "K O P E R A S I"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7275
   Icon            =   "Utama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Utama.frx":27C92
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":267CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2685B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":268E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":269764
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":26A03E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":26A918
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":26B1F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "login2"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "logof2"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rubahpwd_t"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Frm_Anggota_T"
            Object.ToolTipText     =   "Anggota"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "frm_trans_timbang_T"
            Object.ToolTipText     =   "Penimbangan"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "frm_btl_trans_timbang_T"
            Object.ToolTipText     =   "Pembatalan Timbangan"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5670
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/07/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:42"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu fL 
      Caption         =   "&File"
      Begin VB.Menu login 
         Caption         =   "&Login"
      End
      Begin VB.Menu logof 
         Caption         =   "Log &Off"
      End
      Begin VB.Menu grs1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu user 
      Caption         =   "&User"
      Begin VB.Menu User_Baru_M 
         Caption         =   "&Tambah User"
      End
      Begin VB.Menu Form_Hak_Akses_M 
         Caption         =   "&Seting Hak Akses"
      End
      Begin VB.Menu grspw 
         Caption         =   "-"
      End
      Begin VB.Menu rubahpwd_M 
         Caption         =   "&Rubah Password"
      End
   End
   Begin VB.Menu mast 
      Caption         =   "&Master"
      Begin VB.Menu Karyawan_M 
         Caption         =   "&Karyawan"
      End
      Begin VB.Menu Frm_Anggota_M 
         Caption         =   "&Anggota Koperasi"
      End
      Begin VB.Menu grsbsg_X 
         Caption         =   "-"
      End
      Begin VB.Menu tmast_timb_x 
         Caption         =   "Penim&bangan"
         Begin VB.Menu frm_mast_penyusutan_M 
            Caption         =   "&Penyusutan"
         End
         Begin VB.Menu Frm_Harga_Karet_M 
            Caption         =   "&Harga Pabrik Per Kg"
         End
         Begin VB.Menu frm_biaya_adm_M 
            Caption         =   "&Biaya Administrasi"
         End
         Begin VB.Menu frm_biaya_kirim_M 
            Caption         =   "Biaya Pengi&riman"
         End
         Begin VB.Menu Frm_Simpanan_Wajib_M 
            Caption         =   "&Simpanan Wajib Anggota"
         End
      End
   End
   Begin VB.Menu trans 
      Caption         =   "&Transaksi"
      Begin VB.Menu frm_simp_wjb_awal_M 
         Caption         =   "&Jml Saldo Awal Simp Wajib & Sukarela "
      End
      Begin VB.Menu timbang_sl 
         Caption         =   "&Penimbangan Karet"
         Begin VB.Menu frm_trans_timbang_M 
            Caption         =   "&Input Data"
         End
         Begin VB.Menu frm_btl_trans_timbang_M 
            Caption         =   "Pemba&talan"
         End
      End
   End
   Begin VB.Menu lap 
      Caption         =   "&Laporan"
      Begin VB.Menu Frm_sel_Karyawan_M 
         Caption         =   "&Karyawan"
      End
      Begin VB.Menu frm_sel_anggota_M 
         Caption         =   "&Anggota Koperasi"
      End
      Begin VB.Menu frm_sel_penimbangan_M 
         Caption         =   "&Penimbangan Karet"
      End
      Begin VB.Menu frm_sel_tot_wjb_M 
         Caption         =   "&Total Simp Wajib & Sukarela"
      End
   End
End
Attribute VB_Name = "Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim status As String

Public Sub SetAktifMenu(ByVal sql As String)
    
    Dim obj As Object
     Dim a As Long
            
    Dim rec As Recordset
        Set rec = New ADODB.Recordset
            rec.Open sql, kon, adOpenKeyset

    With rec
        If Not .EOF Then
        Do While Not .EOF

               Dim nama_f
               Dim namatol
                    nama_f = !nama_form
                    namatol = nama_f
                    nama_f = nama_f & "_M"
                    namatol = namatol & "_T"
                    
               For Each obj In Me
               
                If TypeOf obj Is Toolbar Then
                Else
                If obj.Name = nama_f Then
                    obj.Enabled = True
                    Exit For
                End If
                End If
                
               Next
                
               
               For a = 1 To 10
                    If UCase(Toolbar1.Buttons.Item(a).Key) = UCase(namatol) Then
                        Toolbar1.Buttons.Item(a).Enabled = True
                        Exit For
                    End If
               Next
                
        .MoveNext
        Loop
        End If

    End With

    rubahpwd_M.Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(1).Enabled = False
    Toolbar1.Buttons.Item(4).Enabled = True
    
End Sub

Private Sub exit_Click()
    End
End Sub

Public Sub enable_menu(ByVal sett As Boolean)
   
   Dim a As Object
   Dim X As Long
        For Each a In Me
        
            If TypeOf a Is Toolbar Then
            Else
            If (UCase(Right(a.Name, 1)) = UCase("M")) Then
                a.Enabled = sett
            End If
            End If
            
        Next
   
        For X = 1 To 10
            If UCase(Right(Toolbar1.Buttons.Item(X).Key, 1)) = UCase("t") Then
                 Toolbar1.Buttons.Item(X).Enabled = sett
            End If
        Next
        
'   adduser_S.Enabled = sett
'   setingakses_S.Enabled = sett
'   rubahpwd_S.Enabled = sett
'
'   kary_S.Enabled = sett
'   anggota_S.Enabled = sett
'   hargaperkilo_S.Enabled = sett
'   simpananwajib_S.Enabled = sett
'
'   timbang_S.Enabled = sett
'   timbang_btl.Enabled = sett
'
'
'   lapkary_S.Enabled = sett
'   lapanggota_S.Enabled = sett
'   laptimbang.Enabled = sett
    
End Sub

Private Sub Form_Hak_Akses_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Form_Hak_Akses
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Form_Hak_Akses
        Frm.Show
    End If

End Sub

Private Sub Frm_Anggota_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Anggota") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Anggota
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Anggota") = False Then Exit Sub
        
        Set Frm = Frm_Anggota
        Frm.Show
    End If


End Sub

Private Sub frm_biaya_adm_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_biaya_adm") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_biaya_adm
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_biaya_adm") = False Then Exit Sub
        
        Set Frm = frm_biaya_adm
        Frm.Show
    End If


End Sub

Private Sub frm_biaya_kirim_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_biaya_kirim") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_biaya_kirim
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_biaya_kirim") = False Then Exit Sub
        
        Set Frm = frm_biaya_kirim
        Frm.Show
    End If


End Sub

Private Sub frm_btl_trans_timbang_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_btl_trans_timbang") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_btl_trans_timbang
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_btl_trans_timbang") = False Then Exit Sub
        
        Set Frm = frm_btl_trans_timbang
        Frm.Show
    End If


End Sub

Private Sub Frm_Harga_Karet_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Harga_Karet") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Harga_Karet
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Harga_Karet") = False Then Exit Sub
        
        Set Frm = Frm_Harga_Karet
        Frm.Show
    End If


End Sub

Private Sub frm_mast_penyusutan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_mast_penyusutan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_mast_penyusutan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_mast_penyusutan") = False Then Exit Sub
        
        Set Frm = frm_mast_penyusutan
        Frm.Show
    End If


End Sub

Private Sub frm_sel_anggota_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_anggota") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_anggota
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_anggota") = False Then Exit Sub
        
        Set Frm = frm_sel_anggota
        Frm.Show
    End If


End Sub

Private Sub Frm_sel_Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_sel_Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Frm_sel_Karyawan
        Frm.Show
    End If


End Sub

Private Sub frm_sel_penimbangan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_penimbangan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_penimbangan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_penimbangan") = False Then Exit Sub
        
        Set Frm = frm_sel_penimbangan
        Frm.Show
    End If


End Sub

Private Sub frm_sel_tot_wjb_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_tot_wjb") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_tot_wjb
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_tot_wjb") = False Then Exit Sub
        
        Set Frm = frm_sel_tot_wjb
        Frm.Show
    End If


End Sub

Private Sub frm_simp_wjb_awal_M_Click()

    Dim konfirmasi As String
        konfirmasi = "Form ini akan memasukkan Simpanan Wajib & Simpanan Sukarela pada saat pertama kali masuk pada program"
        konfirmasi = konfirmasi & Chr(13) & "dan hal itu akan mengakibatkan jml saldo akan kembali sesuai dengan jumlah yang dimasukkan pada form ini"
        konfirmasi = konfirmasi & Chr(13) & "apakah anda akan tetap melanjutkan ?"
        
        If MsgBox(konfirmasi, vbYesNo + vbInformation, "Konfirmasi") = vbNo Then Exit Sub

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_simp_wjb_awal") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_simp_wjb_awal
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_simp_wjb_awal") = False Then Exit Sub
        
        Set Frm = frm_simp_wjb_awal
        Frm.Show
    End If


End Sub

Private Sub Frm_Simpanan_Wajib_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Simpanan_Wajib") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Simpanan_Wajib
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Simpanan_Wajib") = False Then Exit Sub
        
        Set Frm = Frm_Simpanan_Wajib
        Frm.Show
    End If


End Sub

Private Sub frm_trans_timbang_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_trans_timbang") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_trans_timbang
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_trans_timbang") = False Then Exit Sub
        
        Set Frm = frm_trans_timbang
        Frm.Show
    End If


End Sub

Private Sub Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Karyawan
        Frm.Show
    End If


End Sub

Private Sub login_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show

End Sub

Private Sub logof_Click()
    
    If kon.State = adStateClosed Then
            Buka_Koneksi
    End If
    
    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show
    
End Sub

Private Sub MDIForm_Load()
    
    enable_menu False

 status = Buka_Koneksi
 If status = "-2147467259" Then
    
            Dim konfirm As Integer
            Dim Informasi As String
                Informasi = "Koneksi terhadap server tidak berhasil :"
                Informasi = Informasi & vbCrLf & "1. Pastikan server telah hidup dan SQL Server telah dijalankan pada server,atau"
                Informasi = Informasi & vbCrLf & "2. Apabila masih terjadi kegagalan koneksi,periksa nama komputer server,Pastikan nama komputer server tidak berubah"
                Informasi = Informasi & vbCrLf & vbCrLf & "apakah anda ingin menyeting ulang koneksi nama komputer server ?"
        
                konfirm = CInt(MsgBox(Informasi, vbYesNo + vbQuestion, "Konfimasi"))
        
                If konfirm = vbYes Then
        
                    Load Frm_New_Seting
                    Frm_New_Seting.Show
        
                    Unload Me
                    Exit Sub
                Else
                    Unload Me
                    End
                    Exit Sub
                End If

 End If

'    Dim btas As Double
'        btas = Batas
'
'    If btas = 100 Then
'        End
'        Exit Sub
'    Else
'        btas = btas + 1
'        SaveSetting "bts", "bts", "bts", btas
'    End If


    logof.Enabled = False
    Toolbar1.Buttons.Item(2).Enabled = False
    
End Sub



Private Sub rubahpwd_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
                
        Set Frm = Nothing
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
        
    Else
                
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
    End If


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1
            login_Click
        Case 2
            logof_Click
        Case 4
            rubahpwd_M_Click
        Case 6
            Frm_Anggota_M_Click
        Case 7
            frm_trans_timbang_M_Click
        Case 8
            frm_btl_trans_timbang_M_Click
        Case 10
            exit_Click
    End Select
    
End Sub

Private Sub User_Baru_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = User_Baru
        Frm.Show
        
    Else
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = User_Baru
        Frm.Show
    End If
    
End Sub
