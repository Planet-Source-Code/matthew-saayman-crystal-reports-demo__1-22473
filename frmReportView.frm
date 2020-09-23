VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportView 
   Caption         =   "Report Viewer"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8700
   Icon            =   "frmReportView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   8700
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuRepSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuRepclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsReport As Recordset

Public Function RepInit(ByRef RepDesign As Object, _
                          ByVal ADORst As ADODB.Recordset, _
                          ByVal Printer As Boolean, _
                          Optional Formulas As Variant, _
                          Optional CanSave As Boolean = False) As Boolean


Dim iALow As Integer
Dim iAHigh As Integer
Dim i As Integer
On Error GoTo errhandler
    
    RepInit = True
    
    mnuRepSave.Visible = CanSave
    
    If ADORst.RecordCount > 0 Then
        If CanSave = True Then
            Set rsReport = ADORst
            'Create disconnected recordset
            rsReport.ActiveConnection = Nothing
        End If
        
        RepDesign.Database.SetDataSource ADORst, 3, 1
        
        CRViewer1.ReportSource = RepDesign
        ' Assign formulas array to report formulas
        If Not IsNull(Formulas) Then
            iALow = LBound(Formulas)
            iAHigh = UBound(Formulas)
            For i = 1 To iAHigh
                If Formulas(i) <> "" Then
                    RepDesign.FormulaFields(i).Text = Formulas(i)
                End If
            Next
        End If
        
        If Printer = True Then
            ' Print Report
            RepDesign.PrintOut True, 1
            Me.Visible = False
        Else
            Form_Resize
            ' View Report
            CRViewer1.ViewReport
        End If
        
    Else
        MsgBox "There is no data to report on!", vbInformation, "Print Report"
        RepInit = False
    End If
    Set RepDesign = Nothing
    Set ADORst = Nothing
    Exit Function
errhandler:
    On Local Error Resume Next
    RepInit = False
    Set RepDesign = Nothing
    Set ADORst = Nothing
    MsgBox Err.Number & " " & Err.Description
End Function
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = Me.ScaleHeight
    CRViewer1.Width = Me.ScaleWidth
End Sub

Private Sub mnuRepclose_Click()
    Unload Me
End Sub

Private Sub mnuRepSave_Click()
On Error GoTo errhandler

    With CommonDialog1
        .DefaultExt = "RRS"
        .DialogTitle = "Save Report Data"
        .Filename = "*.RRS"
        .Filter = "Report Record Set (*.RRS)"
        .InitDir = App.path
        .CancelError = True
        .ShowSave
    End With
    ' Save disconnected recordset
    rsReport.Save CommonDialog1.Filename
    
    Exit Sub
errhandler:
    ShowError
End Sub
