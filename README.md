<div align="center">

## Autosize a Label Caption


</div>

### Description

This small and very simple sub will format the caption of a Label control if the text is too big to display in the control. The sub will trucate the text and append "..." to the end of the text (indicating to the user that they are not seeing the full text). VB automatically wordwraps the caption of a label if it is too big, however, this results in the caption being truncated only where there is a space. Also, you can see the top of the next line of the caption.

Example

Make and Model: Cadillac

becomes:

Make and Model: Cadillac Eldor...

I find this extremely useful when I don't know the maximum length of the text the label will contain, or if I don't have enough screen real estate to make the Label big enough.

Just pass a label to this sub for formatting.
 
### More Info
 
A label control.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Geoff Temple](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/geoff-temple.md)
**Level**          |Intermediate
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/geoff-temple-autosize-a-label-caption__1-6998/archive/master.zip)





### Source Code

```
'This small and very simple sub will format the
'caption of a Label control if the text is too
'big to display in the control. The sub will
'trucate the text and append "..." to the end
'of the text (indicating to the user that they
'are not seeing the full text). VB automatically
'wordwraps the caption of a label if it is too
'big, however, this results in the caption being
'truncated only where there is a space. Also,
'you can see the top of the next line of the caption.
'Example
'Make and Model: Cadillac
'becomes:
'Make and Model: Cadillac Eldor...
'I find this extremely useful when I don't know the
'maximum length of the text the label will contain,
'or if I don't have enough screen real estate to
'make the Label big enough.
Private Sub AutoSizeCaption(lbl As Label)
  Dim i      As Integer
  Dim iLabelWidth As Integer
  Dim sText    As String
  Const kMore = "..."
  ' store orignal caption and width
  sText = lbl.Caption
  ' numeric or date? Don't format.
  If IsNumeric(lbl.Caption) Or IsDate(lbl.Caption) Then Exit Sub
  iLabelWidth = lbl.Width
  ' allow label to "spring" to it's actual width
  lbl.AutoSize = True
  ' is required width of label < actual width?
  If lbl.Width > iLabelWidth Then
    i = Len(sText) - 1
    Do
      lbl.Caption = Left(sText, i) & kMore
      i = i - 1
    Loop Until (lbl.Width <= iLabelWidth) Or (i = 0)
  End If
Exit_Sub:
  lbl.AutoSize = False
  lbl.Width = iLabelWidth
  Exit Sub
ErrorHandler:
  ' something went wrong ... put everything back
  lbl.Caption = sText
  Resume Exit_Sub
End Sub
```

