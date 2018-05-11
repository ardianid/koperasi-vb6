VERSION 5.00
Begin VB.Form frm_mast_penyusutan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penyusutan"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mast_penyusutan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basah :"
         Height          =   195
         Left            =   615
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kering :"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   555
      End
   End
End
Attribute VB_Name = "frm_mast_penyusutan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Change_Values()
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select * from tb_penyusutan"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        With rs
            
            If Not .EOF Then
                
                
                Text1.Text = IIf(Not IsNull(!kering), !kering, 0)
                Text2.Text = IIf(Not IsNull(!basah), !basah, 0)
                    
            Else
                
                
                Text1.Text = 0
                Text2.Text = 0
                
            End If
            
        End With
        
    
End Sub

Private Sub Command1_Click()
On Error GoTo err_handler
   
    Dim harga_bersih As String
    Dim harga_kotor As String
        
        If Text1.Text = "" Or Text1.Text = 0 Then
            harga_bersih = 0
        Else
            harga_bersih = Replace(Trim(Text1.Text), ",", ".")
        End If
        
        If Text2.Text = "" Or Text2.Text = 0 Then
            harga_kotor = 0
        Else
            harga_kotor = Replace(Trim(Text2.Text), ",", ".")
        End If
        
    Dim sql, sql1 As String
    Dim rs As Recordset
    Dim rs1 As Recordset
        
        sql = "select * from tb_penyusutan"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset
    
    With rs
        
        If Not .EOF Then
            
            sql1 = "update tb_penyusutan set kering=" & harga_bersih & ",basah =" & harga_kotor
            
            Set rs1 = New ADODB.Recordset
                rs1.Open sql1, kon
            
            MsgBox "Data telah dirubah"
        
        Else
            
            sql1 = "insert into tb_penyusutan (kering,basah) values(" & harga_bersih & "," & harga_kotor & ")"
            
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
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    
    With Me
        .Left = Utama.Width / 2 - .Width / 2
        .Top = 1000
    End With
    
    Change_Values
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If


End Sub

Private Sub Text1_GotFocus()
    Call Focus_(Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = "" Then Text1.Text = 0
End Sub

Private Sub Text2_GotFocus()
    Call Focus_(Text2)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text = "" Then Text2.Text = 0
End Sub
