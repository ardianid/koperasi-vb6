VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frm_btl_trans_timbang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pembatalan Transaksi Penimbangan"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_btl_trans_timbang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Keluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   11640
      TabIndex        =   6
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cari"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   7320
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frm_btl_trans_timbang.frx":08CA
      Left            =   2040
      List            =   "frm_btl_trans_timbang.frx":08E3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7320
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_btl_trans_timbang.frx":0902
      Left            =   360
      List            =   "frm_btl_trans_timbang.frx":0912
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   6735
      Left            =   120
      OleObjectBlob   =   "frm_btl_trans_timbang.frx":0940
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   10200
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   6495
   End
End
Attribute VB_Name = "frm_btl_trans_timbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub isi_awal()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from tb_penimbangan where tgl_trans='" & Format(Date, "yyyy/mm/dd") & "' order by id desc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        Set Grid1.DataSource = rs
            Grid1.Refresh
    
End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo err_handler

    Dim sql As String
    Dim rs As Recordset
        
        sql = "select * from tb_penimbangan where "
         
        sql = sql & Combo1.Text
        sql = sql & " " & Combo2.Text
        
        If Combo1.Text = "Tgl_Ambil" Then
        
            If periksa_tanggal(Trim(Text1.Text)) = False Then
                MsgBox "Periksa Kembali tgl pencarian"
                Exit Sub
            End If
            
            If Combo2.Text = "Like" Then
                sql = sql & "'%" & Format(Trim(Text1.Text), "yyyy/mm/dd") & "%'"
            Else
                sql = sql & "'" & Format(Trim(Text1.Text), "yyyy/mm/dd") & "'"
            End If
        Else
            If Combo2.Text = "Like" Then
                sql = sql & "'%" & Trim(Text1.Text) & "%'"
            Else
                sql = sql & "'" & Trim(Text1.Text) & "'"
            End If
        End If
        
        sql = sql & " order by id desc "
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        Set Grid1.DataSource = rs
            Grid1.Refresh
        
        On Error GoTo 0
        Exit Sub
        
err_handler:
        
        MsgBox Error$
        
End Sub

Private Sub Command2_Click()
On Error GoTo err_handler
    
    If Grid1.Row < 0 Then Exit Sub
    
    If MsgBox("Yakin akan dihapus ...", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then Exit Sub
    
    kon.BeginTrans
    
    Dim sql As String
    Dim rs As Recordset
    
    kurangi_simpanan_
    
    sql = "delete from tb_penimbangan where id=" & Grid1.Columns(1).Text
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    kon.CommitTrans
    
    Command3_Click
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
        kon.RollbackTrans
        MsgBox Error$
    
End Sub

Private Sub kurangi_simpanan_()
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "update tb_anggota set jml_wajib=jml_wajib - " & Grid1.Columns(15).Text
        sql = sql & ",jml_sukarela=jml_sukarela - " & Grid1.Columns(16).Text
        sql = sql & " where no_anggota='" & Grid1.Columns(0).Text & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
    
End Sub

Private Sub Command3_Click()
    isi_awal
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
    .Top = 350
End With

isi_awal

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
