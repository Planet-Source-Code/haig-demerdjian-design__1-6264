<div align="center">

## design


</div>

### Description

Creates a pretty neat geometric design!
 
### More Info
 
The number of lines per side you wish to draw. You would enter it into a text box.

After you enter a number into the text box just click on draw and that's it!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Haig Demerdjian](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/haig-demerdjian.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/haig-demerdjian-design__1-6264/archive/master.zip)





### Source Code

```
' Before you start anything you should create: 1 picturebox, 1 textbox, 1 command button. Size the picture box so it is pretty big in size and it should be EXACTLY squared so the design looks nice! To do so make sure the scalewidth and scaleheight in the properties section are equal to each other. Now put the following code into the command button.
Private Sub Command1_Click()
If Text1 <= 0 Then Exit Sub
Picture1.Cls
w = Picture1.ScaleWidth / Text1
h = Picture1.ScaleHeight / Text1
' top to left
For draw = 0 To Text1
  Picture1.Line (0 + (w * draw), 0)-(0, Picture1.ScaleHeight - (h * draw))
Next draw
' left to bottom
For draw = 0 To Text1
  Picture1.Line (0, 0 + (h * draw))-(0 + (w * draw), Picture1.ScaleHeight)
Next draw
' bottom to right
For draw = 0 To Text1
  Picture1.Line (0 + (w * draw), Picture1.ScaleHeight)-(Picture1.ScaleWidth, Picture1.ScaleHeight - (h * draw))
Next draw
' right to top
For draw = 0 To Text1
  Picture1.Line (Picture1.ScaleWidth, 0 + (h * draw))-(0 + (w * draw), 0)
Next draw
End Sub
```

