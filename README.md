<div align="center">

## Insert text


</div>

### Description

This code inserts text at the end of a textbox (or anything with a .text, .selstart, .sellength, and .seltext property) without adding the entire contents of the textbox all over again

it saves a lot of time with long text and opening text files
 
### More Info
 
textcontrol as object,text as string

1 if successful

0 if any error occurs

none i know of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[justin holland](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/justin-holland.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/justin-holland-insert-text__1-899/archive/master.zip)

### API Declarations

nope


### Source Code

```
Function AddText(textcontrol As Object, text2add As String)
  On Error GoTo errhandlr
  tmptxt$ = textcontrol.Text 'just in case of an accident
  textcontrol.SelStart = Len(textcontrol.Text) ' move the "cursor" to the end of the text file
  textcontrol.SelLength = 0 ' highlight nothing (this becomes the selected text)
  textcontrol.SelText = text2add ' set the selected text ot text2add
  AddText = 1
  GoTo quitt ' goto the end of the sub
'error handlers
errhandlr:
  If Err.Number <> 438 Then   'check the error number and restore the
    textcontrol.Text = tmptxt$ 'original text if the control supports it
  End If
  AddText = 0
  GoTo quitt
quitt:
  tmptxt$ = ""
End Function
```

