<div align="center">

## Size into pixels


</div>

### Description

Here is a real simple way to convert the size of vb(twips) into pixels.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SPY\-3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/spy-3.md)
**Level**          |Beginner
**User Rating**    |3.3 (13 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/spy-3-size-into-pixels__1-60111/archive/master.zip)





### Source Code

<b>I found that if you want to convert the vb size(twips) into pixels simply do this<br>
(of course change <font color="green">vbsize</font> to how large the control is ex 495)<br>
<font color="blue">Pixels = </font><font color="green">vbsize </font><font color="blue">/ Screen.TwipsPerPixelX</font> or <font color="blue">Y</font><br>
To turn pixels into vb<br>
(Change pixels into how many pixels ex 33)<br>
<font color="blue">VB = </font><font color="green">pixels</font> <font color="blue">* Screen.TwipsPerPixelX</font> or <font color="blue">Y</font><br>
Examples:<br>
<font color="blue">txtPixelHeight.Text = 495 / Screen.TwipsPerPixelX</font> or <font color="blue">Y</font><br>
<font color="blue">Me.Height = 33 * Screen.TwipsPerPixelX</font> or <font color="blue">Y</font><br>
Also here is a simple function to turn twips to pixels,<br>
<font color="blue">Public Function TwipsToPixel(Twips As Integer, XorY As Boolean) As Integer<br>
If XorY = True Then<br>
TwipsToPixel = Twips / Screen.TwipsPerPixelX<br>
Else<br>
TwipsToPixel = Twips / Screen.TwipsPerPixelY<br>
End If<br>
End Function<br>
</font>
Here is a function to turn Pixels To Twips,<br>
<font color="blue">
Public Function PixelToTwips(Pixels As Integer, XorY As Boolean) As Integer<br>
If XorY = True Then<br>
PixelToTwips = Pixels * Screen.TwipsPerPixelX<br>
Else<br>
PixelToTwips = Pixels * Screen.TwipsPerPixelY<br>
End If<br>
End Function<br>
</font><br>
Hope this helps and please vote.<br>You are free to change or modify this code however you want.

