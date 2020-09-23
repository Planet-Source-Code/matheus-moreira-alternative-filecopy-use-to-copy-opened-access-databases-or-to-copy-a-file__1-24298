<div align="center">

## Alternative FileCopy \- Use to copy opened access databases or to copy a file and make a prog\. bar


</div>

### Description

I made this code because I need to copy an access database with the file open (in use). But, visual basic FileCopy method and windows apis for this pourpose fails in this case with the "File Access Error". So, I made this function that copy the file in blocks. You can alter the block size so the copy can be faster or slower.

Well, thats it. I hope that this code can be useful to anyone!

Ah, the error handle was generated with Ax-Tools CodeSmart 2001, an excelent Add-In for any visual basic programmer! Recommended! :) www.axtools.com
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matheus Moreira](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matheus-moreira.md)
**Level**          |Intermediate
**User Rating**    |4.4 (71 globes from 16 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matheus-moreira-alternative-filecopy-use-to-copy-opened-access-databases-or-to-copy-a-file__1-24298/archive/master.zip)





### Source Code

```
Public Function CopyFile(Source As String, Destiny As String, Optional BlockSize As Long = 32765) As Boolean
    '<EhHeader>
    On Error GoTo CopyFile_Err
    '</EhHeader>
  Dim Pos As Long
  Dim posicao As Long
  Dim pbyte As String
  Dim buffer As Long
  Dim Exist As String
  Dim LenSource As Long
  Dim FFSource As Integer, FFDestiny As Integer
100 buffer = BlockSize
102 posicao = 1
104 Exist = ""
106 Exist = Dir$(Destiny)
108 If Exist <> "" Then Kill Destiny
110 FFSource = FreeFile
112 Open Source For Binary As #FFSource
114 FFDestiny = FreeFile
116 Open Destiny For Binary As #FFDestiny
118 LenSource = LOF(FFSource)
120 For Pos = 1 To LenSource Step buffer
122   If Pos + buffer > LenSource Then buffer = (LenSource - Pos) + 1
124   pbyte = Space$(buffer)
126   Get #FFSource, Pos, pbyte
128   Put #FFDestiny, posicao, pbyte
130   posicao = posicao + buffer
'132   RaiseEvent Progress(Round((((Pos / 100) * 100) / (LenSource / 100)), 2))
'134   DoEvents
  Next
136 Close #FFSource
138 Close #FFDestiny
'140 RaiseEvent CopyComplete
    '<EhFooter>
    Exit Function
CopyFile_Err:
    MsgBox "Um erro inesperado ocorreu!" & vbCrLf & _
        "Por favor anote ou copie (Pressionando a tecla 'Print-Screen' e depois CTRL+V no PAINT) os dados abaixo:" & vbCrLf & _
        "No Erro: " & Err.Number & vbCrLf & _
        "Local: Project1.Form1.CopyFile " & vbCrLf & _
        "Linha: " & Erl & vbCrLf & vbCrLf & _
        "Descrição: " & Err.Description & vbCrLf & vbCrLf & _
        "Operação Cancelada!", vbCritical, "Erro!"
    Screen.MousePointer = vbDefault
    Resume CopyFile_Sai
CopyFile_Sai:
    Exit Function
    '</EhFooter>
End Function
```

