VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_lap_anggota_Detail 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   7965
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_lap_anggota_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New ranggota_detail

Option Explicit

Private Sub Form_Load()
On Error GoTo err_handler

Screen.MousePointer = vbHourglass

Me.Left = 120
Me.Top = 240

Dim sql As String
Dim rs As Recordset
    sql = Mysq
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset

Set Report = New ranggota_detail
    Report.Database.SetDataSource rs, , 1
    
    CRViewer1.Zoom 1

CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

On Error GoTo 0
Exit Sub

err_handler:
    
Screen.MousePointer = vbDefault
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

