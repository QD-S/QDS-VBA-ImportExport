VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InternalUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get TargetObjectName() As String
    TargetObjectName = "ThisWorkbook"
End Property

Public Property Get IsWord() As Boolean
    IsWord = False
End Property

Function ThisAddIn() As Object
    Set ThisAddIn = ThisWorkbook
End Function

Function ActiveTarget() As Object
    Set ActiveTarget = ActiveWorkbook
End Function

Function IsUpdateCodeTarget(Name$) As Boolean
    IsUpdateCodeTarget = Name = TargetObjectName Or ThisWorkbook.Worksheets.Count = 1
End Function
