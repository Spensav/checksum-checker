Attribute VB_Name = "CeksumSementara"
Public Function GetChecksum(FilePath As String) As String
Dim CheckSum(1 To 2) As String
CheckSum(1) = CalcBinary(FilePath, 499, 4500)
CheckSum(2) = CalcBinary(FilePath, 499, 4000)
GetChecksum = CheckSum(1) & CheckSum(2)
End Function
Public Function CalcBinary(ByVal lpFileName As String, ByVal lpByteCount As Long, Optional ByVal StartByte As Long = 0) As String
On Error GoTo Err
Dim Bin() As Byte
Dim ByteSum As Long
Dim i As Long
ReDim Bin(lpByteCount) As Byte
Open lpFileName For Binary As #1
    If StartByte = 0 Then
        Get #1, , Bin
    Else
        Get #1, StartByte, Bin
    End If
Close #1
For i = 0 To lpByteCount
    ByteSum = ByteSum + Bin(i) ^ 2
Next i
CalcBinary = Hex$(ByteSum)
Exit Function
Err:
CalcBinary = "00"
End Function

