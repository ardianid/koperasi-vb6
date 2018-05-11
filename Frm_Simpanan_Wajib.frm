VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form Frm_Simpanan_Wajib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simpanan Wajib Anggota"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Simpanan_Wajib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_bersih 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calculator      =   "Frm_Simpanan_Wajib.frx":08CA
      Caption         =   "Frm_Simpanan_Wajib.frx":08EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frm_Simpanan_Wajib.frx":0956
      Keys            =   "Frm_Simpanan_Wajib.frx":0974
      Spin            =   "Frm_Simpanan_Wajib.frx":09BE
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
      Caption         =   "Jumlah :"
      Height          =   195
      Index           =   0
      Left            =   555
      TabIndex        =   3
      Top             =   480
      Width           =   600
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frm_Simpanan_Wajib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Change_Values()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from Tb_Simpanan_Wajib"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            
            If Not .EOF Then
                
                
                tdb_bersih.Value = IIf(Not IsNull(!jumlah), !jumlah, Null)
    
                    
            Else
                
                
                tdb_bersih.Value = Null
                
            End If
            
        End With
        
    
End Sub


Private Sub Command1_Click()

On Error GoTo err_handler
   
    Dim harga_bersih As Double
        
        If tdb_bersih.ValueIsNull Then
            harga_bersih = 0
        Else
            harga_bersih = Replace(Trim(tdb_bersih.Value), ",", "")
        End If
                
    Dim sql, sql1 As String
    Dim rs As Recordset
    Dim rs1 As Recordset
        
        sql = "select * from Tb_Simpanan_Wajib"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        
        If Not .EOF Then
            
            sql1 = "update Tb_Simpanan_Wajib set jumlah=" & harga_bersih & ""
            
            Set rs1 = New ADODB.Recordset
                rs1.Open sql1, kon
            
            MsgBox "Data telah dirubah"
        
        Else
            
            sql1 = "insert into Tb_Simpanan_Wajib (jumlah) values(" & harga_bersih & ")"
            
            Set rs1 = New ADODB.Recordset
                rs1.Open sql1, kon
            
            MsgBox "Data telah disimpan"
            
        End If
        
    End With
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
        MsgBox Error$

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
    tdb_bersih.SetFocus
End Sub

Private Sub Form_Load()

    With Me
        .Left = Utama.Width / 2 - .Width / 2
        .Top = 1000
    End With
    
    Change_Values
    
End Sub

Private Sub tdb_bersih_GotFocus()
    Call Focus_(tdb_bersih)
End Sub

Private Sub tdb_bersih_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command1.SetFocus
End Sub
