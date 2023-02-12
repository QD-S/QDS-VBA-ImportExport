VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare PtrSafe Function InitializeArrayInternally Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant, Optional ByVal lLow&, Optional ByVal lCount&) As Variant()

Private Const IsCommonVbComponent = True
Private Const IsCommonVbComponentCode$ = "Private Const IsCommonVbComponent = True"
Private Const CommonVbComponentFolderPath$ = ".."
Private Const CodeOfExportFolder$ = "Private Const ExportFolder"

Private Const NewLine = vbCrLf
Private Const DirectorySeparatorChar = "\"
Private Const AltDirectorySeparatorChar = "/"
Private Const VolumeSeparatorChar = ":"

Private Const NotSupportedErr# = 501

Private outputExtensions()
Private outputFolders()
Private outputVBComponentTypes()

Public FileNameFormatText
Public HasTypeFolder As Boolean
Public HasVbaFolder As Boolean

Public Fso As Object

'////////////////////////////////////////////////////Settings////////////////////////////////////////////////////

Private Sub Class_Initialize()
    Set Fso = CreateObject("Scripting.FileSystemObject")
    outputExtensions = Array("bas", "cls", "frm", "bas")
    outputFolders = Array("Modules", "Classes", "Forms", "Objects")
    outputVBComponentTypes = Array(vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document)
    FileNameFormatText = "{FileName}"
    HasTypeFolder = True
    HasVbaFolder = True
End Sub

Private Property Get CodeDirectoryFormat()
    If HasVbaFolder Then CodeDirectoryFormat = FileNameFormatText + ".vba" Else CodeDirectoryFormat = ""
End Property

'////////////////////////////////////////////////////VBComponent////////////////////////////////////////////////////

Function TryGetVBComponent(ByVal VBComponents As VBComponents, Name$, Result As VBComponent) As Boolean
    Dim vbComponent_ As VBComponent
    For Each vbComponent_ In VBComponents
        If vbComponent_.Name = Name Then
            Set Result = vbComponent_
            TryGetVBComponent = True
            Exit Function
        End If
    Next
    TryGetVBComponent = False
End Function

Function GetCode(ByVal VBComponent As VBComponent) As String
    Dim lc_&: lc_ = VBComponent.CodeModule.CountOfLines
    If lc_ > 0 Then
        GetCode = VBComponent.CodeModule.Lines(1, lc_)
    Else
         GetCode = ""
    End If
End Function

Sub ClearCode(ByVal VBComponent As VBComponent)
    VBComponent.CodeModule.DeleteLines 1, VBComponent.CodeModule.CountOfLines
End Sub

Sub UpdateCode(ByVal VBComponent As VBComponent, ByVal Code$)
    ClearCode VBComponent
    VBComponent.CodeModule.AddFromString Code
End Sub

'////////////////////////////////////////////////////Index////////////////////////////////////////////////////

Private Function GetOutputIndexByExtension(ByVal Extension$) As Long
    Dim i&
    For i = LBound(outputExtensions) To UBound(outputExtensions)
        If outputExtensions(i) = Extension Then
            GetOutputIndexByExtension = i
            Exit Function
        End If
    Next
    GetOutputIndexByExtension = -1
End Function

Private Function GetOutputIndex(VBComponent As VBComponent) As Long
    Dim i&, vbComponentType: vbComponentType = VBComponent.Type
    For i = LBound(outputVBComponentTypes) To UBound(outputVBComponentTypes)
        If outputVBComponentTypes(i) = vbComponentType Then
            GetOutputIndex = i
            Exit Function
        End If
    Next
    GetOutputIndex = -1
End Function

'////////////////////////////////////////////////////Arrange////////////////////////////////////////////////////

Sub ArrangeCodeOfAddIn()
    ArrangeCode InternalUtility.ThisAddIn
End Sub

Sub ArrangeCode(Optional ByVal Target)
    If IsMissing(Target) Then Set Target = InternalUtility.ActiveTarget
    ArrangeCodeInternally Target
End Sub

Private Sub ArrangeCodeInternally(ByVal Target)
    ' Remove empty lines before and after code
    Dim vbComponent_ As VBComponent
    Dim startNewLineCount_
    For Each vbComponent_ In Target.VBProject.VBComponents
        Dim code_$: code_ = GetCode(vbComponent_)
        If Len(code_) > 0 Then
            startNewLineCount_ = GetCodeStartNewLineCount(code_, False)
            If startNewLineCount_ > 0 Then vbComponent_.CodeModule.DeleteLines 1, startNewLineCount_
            startNewLineCount_ = GetCodeStartNewLineCount(GetCode(vbComponent_), True)
            If startNewLineCount_ > 0 Then vbComponent_.CodeModule.DeleteLines vbComponent_.CodeModule.CountOfLines - startNewLineCount_ + 1, startNewLineCount_
        End If
    Next
End Sub

Private Function GetCodeStartNewLineCount(ByVal Code$, ByVal IsReverse As Boolean) As Long
    Dim ci_&: ci_ = 0
    Dim nc_&: nc_ = Len(vbCrLf)
    Dim shiftNc_&
    If IsReverse Then
        ci_ = Len(Code) - nc_ + 1
        shiftNc_ = -nc_
    Else
        ci_ = 1
        shiftNc_ = nc_
    End If
    Dim counter_&: counter_ = 0
    Do While Mid(Code, ci_, nc_) = vbCrLf
        ci_ = ci_ + shiftNc_
        counter_ = counter_ + 1
    Loop
    GetCodeStartNewLineCount = counter_
End Function

'////////////////////////////////////////////////////Import////////////////////////////////////////////////////

Sub ImportFilesOfAddIn(Optional ByVal DisableMenuFormOutput As Boolean = False)
    ImportFiles InternalUtility.ThisAddIn, , True, DisableMenuFormOutput, ""
End Sub

Sub ImportFiles(Optional ByVal Target, Optional ByVal ImportPath$, Optional ByVal HasTypeFolders, Optional DisableMenuFormOutput As Boolean = False, Optional ByVal DirectoryFormat)
    
    If MsgBox("Do you import the VB component files?", vbYesNo, "Import") = vbNo Then Exit Sub
    
    If IsMissing(Target) Then Set Target = InternalUtility.ActiveTarget
    Dim HasTypeFolders_ As Boolean: If IsMissing(HasTypeFolders) Then HasTypeFolders_ = VbaUtility.HasTypeFolder Else HasTypeFolders_ = CBool(HasTypeFolders)
    Dim DirectoryFormat_$: If IsMissing(DirectoryFormat) Then DirectoryFormat_ = CodeDirectoryFormat Else DirectoryFormat_ = CStr(DirectoryFormat)
    If ImportPath = "" Then ImportPath = Target.Path
    If DirectoryFormat_ <> "" Then
        Dim codeDirectory_$: codeDirectory_ = Replace(DirectoryFormat_, FileNameFormatText, Target.Name)
        If IsPathRooted(codeDirectory_) Then ImportPath = codeDirectory_ Else ImportPath = Combine(ImportPath, codeDirectory_)
    End If
    
    Dim Fso_ As Object: Set Fso_ = Fso
    If HasTypeFolders_ Then
        Dim i&
        For i = LBound(outputFolders) To UBound(outputFolders)
            ImportFilesInternally Target, GetFiles(Combine(ImportPath, outputFolders(i))), DisableMenuFormOutput
        Next
        For i = LBound(outputFolders) To UBound(outputFolders)
            Dim path_$: path_ = Fso_.GetAbsolutePathName(Combine(ImportPath, CommonVbComponentFolderPath, outputFolders(i)))
            If Fso_.FolderExists(path_) Then
                ImportFilesInternally Target, GetFiles(path_), DisableMenuFormOutput
            End If
        Next
    Else
        ImportFilesInternally Target, GetFiles(ImportPath), DisableMenuFormOutput
    End If

    ArrangeCodeInternally Target
End Sub

Private Sub ImportFilesInternally(ByVal Target, FileInfos, Optional DisableMenuFormOutput As Boolean = False)
    Dim fileInfo_
    For Each fileInfo_ In FileInfos
        Dim vbComponents_ As VBComponents: Set vbComponents_ = Target.VBProject.VBComponents
        Dim extension_$: extension_ = GetExtension(fileInfo_.Name)
        Dim index_&: index_ = GetOutputIndexByExtension(extension_)
        If index_ >= 0 Then
            Dim vbComponent_ As VBComponent
            Dim fileNameWithoutExtension_$: fileNameWithoutExtension_ = Fso.GetBaseName(fileInfo_.Name)
            If TryGetVBComponent(vbComponents_, fileNameWithoutExtension_, vbComponent_) Then
                Dim fileCode_$: fileCode_ = GetCodeFromFile(fileInfo_.Path, DefaultCharset)
                Dim fileCode2_$: fileCode2_ = GetCodeFromFile(fileInfo_.Path)
                Dim code_$: code_ = GetCode(vbComponent_) + NewLine
                If fileCode_ = code_ Or fileCode2_ = code_ Then
                    GoTo ContinueLabel
                End If
                
                Dim vbComponentType: vbComponentType = vbComponent_.Type
                Dim hasEnum_ As Boolean
                If InternalUtility.IsWord Then
                    Dim enumCount_&: enumCount_ = IndexOfCount(fileCode_, " Enum ") + IndexOfCount(fileCode_, NewLine + "Enum ")
                    hasEnum_ = enumCount_ > 0
                    If hasEnum_ Then hasEnum_ = IndexOfCount(fileCode_, "Private Enum ") < enumCount_
                Else
                    hasEnum_ = False
                End If
                If hasEnum_ Then
                    UpdateCode vbComponent_, GetCodeFromFile(fileInfo_.Path)
                    GoTo ContinueLabel
                ElseIf vbComponentType = vbext_ct_Document And InternalUtility.IsUpdateCodeTarget(fileNameWithoutExtension_) Then
                    Dim vbc_ As VBComponent: Set vbc_ = vbComponents_.Import(fileInfo_.Path)
                    UpdateCode vbComponent_, GetCode(vbc_)
                    vbComponents_.Remove vbc_
                    GoTo ContinueLabel
                ElseIf DisableMenuFormOutput And fileNameWithoutExtension_ = "MenuForm" Then
                    MsgBox "MenuForm code is changed. You have to run 'Utility_.ImportFilesOfAddIn' directly."
                    GoTo ContinueLabel
                ElseIf vbComponentType <> vbext_ct_Document And vbComponentType <> vbext_ct_MSForm Then
                    vbComponent_.Name = fileNameWithoutExtension_ & "Z"
                End If
                vbComponents_.Remove vbComponent_
            End If
            DoEvents
            vbComponents_.Import fileInfo_.Path
ContinueLabel:
        End If
    Next
End Sub

Function GetCodeFromFile(ByVal Path$, Optional ByVal Charset$)
    Dim text_$: text_ = Read(Path, Charset)
    Const AttributeCode_$ = "Attribute"
    Dim ci_&: ci_ = 1
    Dim pi_&: pi_ = InStr(ci_, text_, AttributeCode_)
    Do While True
        ci_ = InStr(pi_, text_, vbCrLf)
        If ci_ <= 0 Then
            Exit Do
        End If
        Dim line_$: line_ = Mid(text_, pi_, ci_)
        If Not StartsWith(line_, AttributeCode_) Then
            Exit Do
        End If
        pi_ = ci_ + Len(vbCrLf)
    Loop
    GetCodeFromFile = Mid(text_, pi_)
End Function

Function GetCodeByLineCountFromFile(ByVal Path$, ByVal LineCount&, ByVal CodeLength&, Optional ByVal Charset$)
    Dim text_$: text_ = Read(Path, Charset)
    Dim ci&: ci = LastIndexOf(text_, NewLine, , , LineCount)
    GetCodeByLineCountFromFile = Left(Mid(text_, ci), CodeLength)
End Function

Function LastIndexOf(ByVal Target$, ByVal Value$, Optional ByVal StartIndex& = -1, Optional ByVal Count& = -1, Optional ByVal RepeatCount& = 1)

    If StartIndex = -1 Then StartIndex = Len(Target)
    If Count = -1 Then Count = Len(Target)

    Dim rangeStart_&: rangeStart_ = StartIndex - Count + 1
    Dim i&, ci&
    ci = StartIndex
    For i = 1 To RepeatCount
        ci = InStrRev(Target, Value, ci)
        If ci = 0 Or ci < rangeStart_ Then
            LastIndexOf = -1
            Exit Function
        End If
    Next
    LastIndexOf = ci

End Function

'////////////////////////////////////////////////////Export////////////////////////////////////////////////////

Sub ExportFilesOfAddIn()
    ExportFiles InternalUtility.ThisAddIn, , True, ""
End Sub

Sub ExportFiles(Optional ByVal Target, Optional ByVal ExportPath$, Optional ByVal HasTypeFolders, Optional ByVal DirectoryFormat)
    If IsMissing(Target) Then Set Target = InternalUtility.ActiveTarget
    Dim HasTypeFolders_ As Boolean: If IsMissing(HasTypeFolders) Then HasTypeFolders_ = VbaUtility.HasTypeFolder Else HasTypeFolders_ = CBool(HasTypeFolders)
    Dim DirectoryFormat_$: If IsMissing(DirectoryFormat) Then DirectoryFormat_ = CodeDirectoryFormat Else DirectoryFormat_ = CStr(DirectoryFormat)
    If ExportPath = "" Then ExportPath = Target.Path
    If DirectoryFormat_ <> "" Then
        Dim codeDirectory_$: codeDirectory_ = Replace(DirectoryFormat_, FileNameFormatText, Target.Name)
        If IsPathRooted(codeDirectory_) Then ExportPath = codeDirectory_ Else ExportPath = Combine(ExportPath, codeDirectory_)
    End If
    
    Dim vbComponent_ As VBComponent
    Dim folder_$
    For Each vbComponent_ In Target.VBProject.VBComponents
        Dim index_&: index_ = GetOutputIndex(vbComponent_)
        If index_ <> -1 Then
        
            Dim code_$: code_ = GetCode(vbComponent_)
            Dim isTarget_ As Boolean
            If vbComponent_.Type = vbext_ct_Document Then
                Dim trimedCode$: trimedCode = Trim(code_)
                If trimedCode = "" Or trimedCode = "Option Explicit" Then
                    GoTo ContinueLabel
                End If
            End If
            
            Dim directoryPath_$: directoryPath_ = ExportPath
            Dim moduleExportDirectory$
            If TryGetExportDirectory(vbComponent_, moduleExportDirectory) Then directoryPath_ = Combine(directoryPath_, moduleExportDirectory)
            If HasTypeFolders_ Then directoryPath_ = Combine(directoryPath_, outputFolders(index_))
            
            Dim componentName$: componentName = vbComponent_.Name
            Dim fileName_$: fileName_ = componentName & "." & outputExtensions(index_)
            If HasTypeFolders_ Then
                Dim underlineIndex&: underlineIndex = InStr(1, componentName, "_")
                If underlineIndex > 0 And underlineIndex < Len(componentName) Then directoryPath_ = Combine(directoryPath_, Left(fileName_, underlineIndex - 1))
            End If
            directoryPath_ = GetFullPath(directoryPath_)
            CreateDirectory directoryPath_
            
            Dim filePath_$: filePath_ = Combine(directoryPath_, fileName_)
            If ExistsAsFile(filePath_) Then
                code_ = code_ + NewLine
                Dim fileCode_$: fileCode_ = GetCodeFromFile(filePath_, DefaultCharset)
                Dim fileCode2_$: fileCode2_ = GetCodeFromFile(filePath_)
                If fileCode_ = code_ Then
                    GoTo ContinueLabel
                ElseIf fileCode2_ = code_ Then
                    GoTo ContinueLabel
                End If
            End If
            
            vbComponent_.Export filePath_
        End If
ContinueLabel:
    Next
End Sub

Private Function TryGetExportDirectory(VBComponent As VBComponent, Result$) As Boolean
    Dim code_$: code_ = VBComponent.CodeModule.Lines(1, GetMax(VBComponent.CodeModule.CountOfDeclarationLines, 1))
    If Len(code_) > 0 Then
        Dim currentIndex&, nextIndex&: currentIndex = 1
        Dim newLineLength&: newLineLength = Len(NewLine)
        Dim cl_&: cl_ = Len(code_)
        Do While currentIndex <= cl_
            nextIndex = InStr(currentIndex, code_, NewLine)
            If nextIndex = 0 Then nextIndex = cl_ + 1
            Dim line_$: line_ = Mid(code_, currentIndex, nextIndex - currentIndex)
            currentIndex = nextIndex + newLineLength
            If StartsWith(line_, CodeOfExportFolder) Then
                Dim endDoubleQuoteIndex&: endDoubleQuoteIndex = InStrRev(line_, """") - 1
                Dim startDoubleQuoteIndex&: startDoubleQuoteIndex = InStrRev(line_, """", endDoubleQuoteIndex - 1) + 1
                Result = Cut(line_, startDoubleQuoteIndex, endDoubleQuoteIndex)
                TryGetExportDirectory = True
                Exit Function
            ElseIf line_ = IsCommonVbComponentCode Then
                Result = CommonVbComponentFolderPath
                TryGetExportDirectory = True
                Exit Function
            End If
        Loop
    End If
    
    Result = Empty
    TryGetExportDirectory = False
End Function

'////////////////////////////////////////////////////Array////////////////////////////////////////////////////

Function InitializeArray(Optional ByVal VbVarType As VbVarType = vbVariant, Optional ByVal StartIndex& = 0, Optional ByVal Count& = 0)
    If Count = 0 Then
        InitializeArray = InitializeArrayInternally(VbVarType, StartIndex, 0)
    Else
        Dim result_(): ReDim result_(StartIndex To StartIndex + Count - 1)
        InitializeArray = result_
    End If
End Function

Function Create(ByVal ArrayLength&, Optional ByVal DefaultValue, Optional ByVal ArrayStart)
    Dim as_&: If IsMissing(ArrayStart) Then as_ = 0 Else as_ = ArrayStart
    Dim result_()
    If Not IsMissing(DefaultValue) Then
       result_ = InitializeArray(VarType(DefaultValue), as_, ArrayLength)
        Dim i&
        For i = LBound(result_) To UBound(result_)
            result_(i) = DefaultValue
        Next
    Else
        result_ = InitializeArray(vbVariant, as_, ArrayLength)
    End If
    Create = result_
End Function

Function CreateArray(Target1, Optional Target2, Optional Target3, Optional Target4)
    Dim targetCount#
    If Not IsMissing(Target4) Then
        targetCount = 4
    ElseIf Not IsMissing(Target3) Then
        targetCount = 3
    ElseIf Not IsMissing(Target2) Then
        targetCount = 2
    Else
        targetCount = 1
    End If
    Dim result_(): ReDim result_(0 To targetCount - 1)
    result_(0) = Target1
    If Not IsMissing(Target2) Then result_(1) = Target2
    If Not IsMissing(Target3) Then result_(2) = Target3
    If Not IsMissing(Target4) Then result_(3) = Target4
    CreateArray = result_
End Function

Function ToArray(ByVal Target, Optional ByVal TargetDimension& = 1, Optional ByVal IsReferenceType As Boolean)

    Dim i&, result_()
    If TypeOf Target Is Collection Then
        result_ = Create(Target.Count)
        Dim obj_
        For Each obj_ In Target
            If VarType(obj_) = 9 Or IsReferenceType Then
                Set result_(i) = obj_
            Else
                result_(i) = obj_
            End If
            i = i + 1
        Next
    Else
        Dim ValueArrays: ValueArrays = Target
        If Not IsArray(ValueArrays) Then
            result_ = CreateArray(ValueArrays)
        ElseIf TargetDimension = 1 Then
            result_ = ValueArrays
        ElseIf TargetDimension = 2 Then
            Dim j&, ci_&
            ci_ = LBound(ValueArrays, 1)
            result_ = Create((UBound(ValueArrays, 1) - LBound(ValueArrays, 1) + 1) * (UBound(ValueArrays, 2) - LBound(ValueArrays, 2) + 1), ArrayStart:=ci_)
            For i = LBound(ValueArrays, 1) To UBound(ValueArrays, 1)
                For j = LBound(ValueArrays, 2) To UBound(ValueArrays, 2)
                    result_(ci_) = ValueArrays(i, j)
                    ci_ = ci_ + 1
                Next
            Next
        Else
            Err.Raise NotSupportedErr
        End If
    End If
    ToArray = result_
End Function

'////////////////////////////////////////////////////File////////////////////////////////////////////////////

Function GetFiles(ByVal DirectoryPath$, Optional ByVal FileNameFlter$, Optional ByVal IncludeSubDirectories As Boolean = True)
    Dim checkFileName_$: checkFileName_ = StripStart(FileNameFlter, "*.")
    Dim isExtensionFilter As Boolean: isExtensionFilter = checkFileName_ <> FileNameFlter

    Dim file_
    Dim rc_ As Collection: Set rc_ = New Collection
    Dim directory_
    Dim subFolderResult
    If Exists(DirectoryPath) Then
        If IncludeSubDirectories Then
            For Each directory_ In Fso.GetFolder(DirectoryPath).SubFolders
                Dim subFolderResults: subFolderResults = GetFiles(CStr(directory_), FileNameFlter)
                For Each subFolderResult In subFolderResults
                    rc_.Add subFolderResult
                Next
            Next
        End If
    
        For Each file_ In Fso.GetFolder(DirectoryPath).Files
            If isExtensionFilter Then
                If EndsWith(file_.Name, checkFileName_) Then
                    rc_.Add file_
                End If
            Else
                rc_.Add file_
            End If
        Next
    End If

    GetFiles = ToArray(rc_, IsReferenceType:=True)
End Function

Function Read(filePath$, Optional ByVal Charset$)
    Dim result_$
    
    If Charset = "" Then Charset = "UTF-8"
    
    With CreateObject("ADODB.Stream")
        .Charset = Charset
        .Open
        .LoadFromFile filePath
        result_ = .ReadText
        .Close
    End With
    
    Read = result_
End Function

Function Exists(ByVal Path$)
    Dim result_ As Boolean
    If Fso.FileExists(Path) Then
        result_ = True
    ElseIf Fso.FolderExists(Path) Then
        result_ = True
    Else
        result_ = False
    End If
    Exists = result_
End Function

Function ExistsAsFile(ByVal Path$) As Boolean
    ExistsAsFile = Fso.FileExists(Path)
End Function

'////////////////////////////////////////////////////Path////////////////////////////////////////////////////

Function GetDirectory(ByVal Target$) As String
    Dim StartIndex&
    If EndsWith(Target, DirectorySeparatorChar) Then
        StartIndex = Len(Target) - 1
    Else
        StartIndex = Len(Target)
    End If
    GetDirectory = Left(Target, InStrRev(Target, DirectorySeparatorChar, StartIndex))
End Function

Sub CreateDirectory(ByVal Target$)
    Dim parentDirectory$: parentDirectory = GetDirectory(Target)

    If Not Fso.FolderExists(parentDirectory) Then
        CreateDirectory parentDirectory
    End If

    If Not Fso.FolderExists(Target) Then
        MkDir Target
    End If
End Sub

Function GetExtension(ByVal Target$, Optional ByVal WithDot As Boolean) As String
    Dim dotIndex&: dotIndex = InStrRev(Target, ".")
    If dotIndex = -1 Then
        GetExtension = ""
    Else
        Dim dotShift&: If WithDot Then dotShift = 1
        GetExtension = Right(Target, Len(Target) - dotIndex + dotShift)
    End If
End Function

Private Function IsPathRooted(ByVal Target$) As Boolean
    Dim length_&: length_ = Len(Target)
    If length_ > 0 Then
        Dim char_: char_ = Left(Target, 1)
        If char_ = DirectorySeparatorChar Then
            IsPathRooted = True
        ElseIf char_ = AltDirectorySeparatorChar Then
            IsPathRooted = True
        ElseIf length_ > 1 Then
            IsPathRooted = Mid(Target, 2, 1) = VolumeSeparatorChar
        Else
            IsPathRooted = False
        End If
    Else
        IsPathRooted = False
    End If
End Function

Function GetFullPath(ByVal Target$) As String
    GetFullPath = Fso.GetAbsolutePathName(Target)
End Function

Private Function Combine(ParamArray Values()) As String
    Dim newValues_: ReDim newValues_(LBound(Values) To UBound(Values))
    Dim i&
    For i = LBound(Values) To UBound(Values)
        newValues_(i) = StripEnd(Values(i), DirectorySeparatorChar)
    Next
    Combine = Join(newValues_, DirectorySeparatorChar)
End Function

'////////////////////////////////////////////////////String////////////////////////////////////////////////////

Function IndexOfCount(ByVal Target$, ByVal Value$) As Long
    Dim counter_&: counter_ = 0
    Dim ci_&: ci_ = 1
    Do While True
        ci_ = InStr(ci_, Target, Value)
        If ci_ <= 0 Then Exit Do
        counter_ = counter_ + 1
    Loop
    IndexOfCount = counter_
End Function

Function StartsWith(ByVal Target$, ByVal Value$) As Boolean
    If Len(Target) >= Len(Value) Then
        If Left(Target, Len(Value)) = Value Then
            StartsWith = True
            Exit Function
        End If
    End If

    StartsWith = False
End Function

Function EndsWith(ByVal Target$, ByVal Value$) As Boolean
    If Len(Target) >= Len(Value) Then
        If Right(Target, Len(Value)) = Value Then
            EndsWith = True
            Exit Function
        End If
    End If

    EndsWith = False
End Function

Function StripStart(ByVal Target$, ByVal Value) As String
    Dim rv_$
    If StartsWith(Target, Value) Then
        StripStart = Right(Target, Len(Target) - Len(Value))
    Else
        StripStart = Target
    End If
End Function

Function StripEnd(ByVal Target$, ByVal Value) As String
    Dim rv_$
    If EndsWith(Target, Value) Then
        StripEnd = Left(Target, Len(Target) - Len(Value))
    Else
        StripEnd = Target
    End If
End Function

Function Cut(ByVal Target$, ByVal StartIndex&, ByVal EndIndex&) As String
    Cut = Mid(Target, StartIndex, EndIndex - StartIndex + 1)
End Function

'////////////////////////////////////////////////////Math////////////////////////////////////////////////////

Function GetMax(ByVal Value1, ByVal Value2)
    If Value1 < Value2 Then
        GetMax = Value2
    Else
        GetMax = Value1
    End If
End Function


