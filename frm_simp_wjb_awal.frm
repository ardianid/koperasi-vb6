VERSION 5.00
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frm_simp_wjb_awal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jml Simpanan Wajib & Sukarela Awal Anggota"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_simp_wjb_awal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Daftar 
      Height          =   3855
      Left            =   -6120
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   6800
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_simp_wjb_awal.frx":08CA
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_simp_wjb_awal.frx":08E6
      Childs          =   "frm_simp_wjb_awal.frx":0992
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
         TabIndex        =   10
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Txt_Cr_Daftar 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Daftar 
         Height          =   2775
         Left            =   240
         OleObjectBlob   =   "frm_simp_wjb_awal.frx":09AE
         TabIndex        =   11
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Index           =   41
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Anggota"
         Height          =   195
         Index           =   40
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton ckeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton csimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   6360
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid grid1 
      Height          =   5535
      Left            =   120
      OleObjectBlob   =   "frm_simp_wjb_awal.frx":3932
      TabIndex        =   4
      Top             =   600
      Width           =   9015
   End
   Begin VB.CommandButton ctampil 
      Caption         =   "&Tampil"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cbrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Anggota :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frm_simp_wjb_awal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrgrid As New XArrayDB
Dim Moving As Boolean
Dim yold, xold As Long


Private Sub kosong_grid()
    
    arrgrid.ReDim 0, 0, 0, grid1.Columns.Count
    arrgrid.ReDim 1, 1, 1, grid1.Columns.Count
        grid1.ReBind
        grid1.Refresh
    
End Sub

Private Sub cbrowse_Click()

With TDB_Daftar

If .Visible = False Then
    
    .Left = cbrowse.Left + cbrowse.Width / 2 - .Width / 2
    .Top = cbrowse.Top + cbrowse.Height + 15
    
        
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

Private Sub ckeluar_Click()
    Unload Me
End Sub

Private Sub csimpan_Click()
On Error GoTo err_handler

    If arrgrid.UpperBound(1) = 1 And arrgrid(1, 1) = Empty Then Exit Sub
    
    If MsgBox("Yakin sudah benar ...", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
        Exit Sub
    End If
    
    kon.BeginTrans
    
    Dim a As Long
    Dim sql As String
    Dim rs As Recordset
        For a = 1 To arrgrid.UpperBound(1)
            
            sql = "update tb_anggota set jml_wajib=" & arrgrid(a, 3) & ",jml_sukarela=" & arrgrid(a, 4)
            sql = sql & " where no_anggota='" & arrgrid(a, 1) & "'"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
            
            
        Next
    
    kon.CommitTrans
    MsgBox "Data telah disimpan"
    Exit Sub
        
err_handler:
    
    kon.RollbackTrans
    MsgBox Error$
   
End Sub

Private Sub ctampil_Click()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim a As Long
        a = 1
    
    Dim no, nama As String
    
    kosong_grid
    
    If UCase(Text1.Text) = UCase("semua") Then
        sql = "select * from tb_anggota order by nama asc"
    Else
        sql = "select * from tb_anggota"
        sql = sql & " where no_anggota ='" & Trim(Text1.Text) & "' order by nama asc"
    End If
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        If Not .EOF Then
            
            Do While Not .EOF
                arrgrid.ReDim 1, a, 0, grid1.Columns.Count
                    grid1.ReBind
                    grid1.Refresh
                
                no = IIf(Not IsNull(!no_anggota), !no_anggota, "")
                nama = IIf(Not IsNull(!nama), !nama, "")
            
            arrgrid(a, 0) = a
            arrgrid(a, 1) = no
            arrgrid(a, 2) = nama
            arrgrid(a, 3) = 0
            arrgrid(a, 4) = 0
            
            DoEvents
            
            a = a + 1
            .MoveNext
            Loop
            
            grid1.ReBind
            grid1.Refresh
            
            grid1.MoveFirst
            
        End If
    End With
    
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Text1.SetFocus
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

    Text1.Text = "Semua"
    grid1.Array = arrgrid
    
    kosong_grid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If

End Sub

Private Sub grid_daftar_DblClick()
    
    If Grid_Daftar.Row < 0 Then Exit Sub
    
    Text1.Text = Grid_Daftar.Columns(0).Text
    
    TDB_Daftar.Visible = False
    ctampil.SetFocus
    
End Sub

Private Sub grid_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grid_daftar_DblClick
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: Text1.SetFocus
End Sub

Private Sub grid1_AfterColUpdate(ByVal ColIndex As Integer)
    
    If ColIndex = 3 Or ColIndex = 4 Then
        arrgrid(grid1.Bookmark, ColIndex) = grid1.Columns(ColIndex).Text
        DoEvents
    End If
    
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

Private Sub Text1_GotFocus()
    Call Focus_(Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ctampil.SetFocus
    If KeyCode = vbKeyF3 Then cbrowse_Click
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = "" Then Text1.Text = "Semua"
End Sub

Private Sub Txt_Cr_Daftar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Daftar.SetFocus
    If KeyCode = vbKeyEscape Then TDB_Daftar.Visible = False: Text1.SetFocus
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
