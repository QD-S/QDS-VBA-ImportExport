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
    TargetObjectName = "ThisDocument"
End Property

Public Property Get IsWord() As Boolean
    IsWord = True
End Property

Function ThisAddIn() As Object
    Set ThisAddIn = ThisDocument
End Function

Function ActiveTarget() As Object
    Set ActiveTarget = ActiveDocument
End Function

Function AddMonths(Target As Date, Months&) As Date
    AddMonths = DateAdd("m", Months, Target)
End Function

Function IsUpdateCodeTarget(Name$) As Boolean
    IsUpdateCodeTarget = Name = InternalUtility.TargetObjectName
End Function

