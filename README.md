<div align="center">

## TV Lines Image Filter

<img src="PIC20017211521354242.gif">
</div>

### Description

This code puts lines over a picture box's picture. You can set it's opacity, and it's direction.
 
### More Info
 
PictBox, the picture box to manipulate. optional Opacity. This controls how transparent/solid the lines are. The value works best with a value 1-100. And direction...wich sets the lines horizontal(1) or vertical(2).

For this to work, the picturebox must have AutoRedraw set, and it's scalemode must be pixels.

if it's a very big image, it might be a tiny bit slow.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jared Collums](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jared-collums.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jared-collums-tv-lines-image-filter__1-25300/archive/master.zip)

### API Declarations

```
Public Declare Function SetPixel Lib "GDI32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
```


### Source Code

```
Public Sub TVLines(PictBox As PictureBox, Optional Direction As Integer, Optional Opacity As Long)
Dim i As Long, k As Long, r As Long, g As Long, b As Long, pixel As Long, pix As Long
If IsMissing(Opacity) Then Opacity = 25
If IsMissing(Direction) Then Direction = 1
Opacity = Opacity * 2.55
Opacity = Round(Opacity)
For k = 0 To PictBox.ScaleHeight - 1
 For i = 0 To PictBox.ScaleWidth - 1
 'get current pixel
 pixel = GetPixel(PictBox.HDC, i, k)
 'get rgb values of the pixel
 r = TakeRGB(pixel, 0)
 g = TakeRGB(pixel, 1)
 b = TakeRGB(pixel, 2)
 'the code alternates lightness/darkness each line
 If Direction = 1 Then
 pix = k
 Else
 pix = i
 End If
 If pix / 2 = Int(pix / 2) Then
 r = IIf(r - Opacity < 0, 0, r - Opacity)
 g = IIf(g - Opacity < 0, 0, g - Opacity)
 b = IIf(b - Opacity < 0, 0, b - Opacity)
 Else
 r = IIf(r + Opacity > 255, 255, r + Opacity)
 g = IIf(g + Opacity > 255, 255, g + Opacity)
 b = IIf(b + Opacity > 255, 255, b + Opacity)
 End If
 'set new pixel
 SetPixel PictBox.HDC, i, k, RGB(r, g, b)
 Next i
 PictBox.Refresh
Next k
PictBox.Refresh
End Sub
'just a function to get rgb values of a pixel
'I borrowed it from Jongmin Baek's Drawer (an exellect program, btw)
Function TakeRGB(Colors As Long, Index As Long) As Long
IndexColor = Colors
Red = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Red) / 256
Green = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Green) / 256
Blue = IndexColor
If Index = 0 Then TakeRGB = Red
If Index = 1 Then TakeRGB = Green
If Index = 2 Then TakeRGB = Blue
End Function
```

