Attribute VB_Name = "getRTD"
Option Explicit
Option Compare Text
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       3rd April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Dim mCollection As VBA.Collection
Dim oShell As IWshRuntimeLibrary.WshShell

Function GetRTDServersUsedByActiveWorkbook()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       3rd April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const strMethodName As String = "getRTD.GetRTDServersUsedByActiveWorkbook "
Const strRTDSignature = "=RTD("""
'Variables
Dim oWorkbook As Excel.Workbook
Dim oWorkSheet As Excel.Worksheet
Dim oCell As Excel.Range
Dim strRTDProgID As String
Dim lngCount As Long
Dim strRTDServerFileLocation As String
Dim strClassIDFromProgID As String
Dim strCodeBase As String

Set oWorkbook = ThisWorkbook

For Each oWorkSheet In oWorkbook.Worksheets
    For Each oCell In oWorkSheet.UsedRange.Cells
        If StrComp(strRTDSignature, Left(oCell.Formula, Len(strRTDSignature)), vbTextCompare) = 0 Then
            strRTDProgID = Mid(oCell.Formula, (Len(strRTDSignature) + 1), InStr(Len(strRTDSignature) + 1, oCell.Formula, """", vbTextCompare) - (Len(strRTDSignature) + 1))
            'Debug.Print "[" & oCell.Address & "]" & oCell.Formula & "{" & strRTDProgID & "}"
            AddToCollection (strRTDProgID)
        End If
    Next
Next

'We have a collection of RTD ProgIDs now...
Set oShell = New IWshRuntimeLibrary.WshShell

For lngCount = 1 To mCollection.count
    strRTDServerFileLocation = GetCodeBase(mCollection(lngCount))
    If Len(strRTDServerFileLocation) <> 0 Then
        Debug.Print "RTD Server Codebase located at:" & GetCodeBase(mCollection(lngCount))
    End If
Next

End Function

Function GetCodeBase(ProgID As String) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       3rd April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const strMethodName As String = "getRTD.GetCodeBase "
Const strRegRoot As String = "HKEY_CLASSES_ROOT\"
Const strRegClsId As String = "CLSID\"
Const strRegInprocServerCodeBase As String = "\InprocServer32\Codebase"
'Variables
Dim strCodeBase As String
Dim strClassIDFromProgID As String

'HKEY_CLASSES_ROOT\alignmentsystems.node2rtd2\CLSID
'HKEY_CLASSES_ROOT\CLSID\{B856EB46-BCA3-4A44-9DC9-8995C37680CB}

'-----------------------------------
On Error GoTo ErrHandler1
strClassIDFromProgID = oShell.RegRead(strRegRoot & ProgID & "\" & strRegClsId)
'Debug.Print strClassIDFromProgID

On Error GoTo 0


'-----------------------------------
On Error GoTo ErrHandler2
strCodeBase = oShell.RegRead(strRegRoot & strRegClsId & strClassIDFromProgID & strRegInprocServerCodeBase)
'Debug.Print strCodeBase

GetCodeBase = strCodeBase

Exit Function
ErrHandler1:
'We failed to get the HKEY_CLASSES_ROOT\progid\CLSID
GetCodeBase = ""
Exit Function

ErrHandler2:
'We failed to get the HKEY_CLASSES_ROOT\CLSID
GetCodeBase = ""
Exit Function



End Function

Function AddToCollection(RTDStringOfProgID As String)
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       3rd April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const strMethodName As String = "getRTD.RTDStringOfProgID "

If mCollection Is Nothing Then
    Set mCollection = New Collection
End If

On Error Resume Next
mCollection.Add RTDStringOfProgID, RTDStringOfProgID
On Error GoTo 0

End Function
