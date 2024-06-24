Attribute VB_Name = "modColourBar"
Option Explicit

Public Sub ColourBar(pb As PictureBox, _
                     ByVal Low As Single, _
                     ByVal High As Single, _
                     ByVal Value As Single)

      Dim intR As Integer
      Dim intG As Integer
      Dim intB As Integer
      Dim lngC As Long
      Dim sngX As Single
      Dim sngInc As Single
      Dim sngBtmBar As Single
      Dim sngTopBar As Single

59540 pb.Cls
59550 pb.DrawWidth = 1

59560 intB = 0
59570 intG = 0
59580 intR = 255

59590 sngBtmBar = pb.height * 0.666
59600 sngTopBar = pb.height * 0.333

59610 sngInc = pb.ScaleWidth / (255 * 4)
59620 sngX = 0

59630 For intG = 0 To 255
59640   lngC = rgb(intR, intG, intB)
59650   pb.Line (sngX, sngTopBar)-(sngX, sngBtmBar), lngC
59660   sngX = sngX + sngInc
59670 Next

59680 intG = 255
59690 For intR = 255 To 0 Step -1
59700   lngC = rgb(intR, intG, intB)
59710   pb.Line (sngX, sngTopBar)-(sngX, sngBtmBar), lngC
59720   sngX = sngX + sngInc
59730 Next

59740 For intR = 0 To 255
59750   lngC = rgb(intR, intG, intB)
59760   pb.Line (sngX, sngTopBar)-(sngX, sngBtmBar), lngC
59770   sngX = sngX + sngInc
59780 Next

59790 For intG = 255 To 0 Step -1
59800   lngC = rgb(intR, intG, intB)
59810   pb.Line (sngX, sngTopBar)-(sngX, sngBtmBar), lngC
59820   sngX = sngX + sngInc
59830 Next

59840 pb.DrawWidth = 2
59850 If Value <> 0 And Low <> 0 And High <> 0 Then
59860   sngX = (Value - Low) / (High - Low)
59870   pb.Line (sngX * 1024, 0)-(sngX * 1024, pb.height), vbBlack
59880 End If

End Sub
Public Sub ColourBarVert(pb As PictureBox)

      Dim intR As Integer
      Dim intG As Integer
      Dim intB As Integer
      Dim lngC As Long
      Dim sngY As Single
      Dim sngInc As Single

59890 intB = 0
59900 intG = 0
59910 intR = 255

59920 sngInc = pb.height / (255 * 4)
59930 sngY = 0

59940 For intG = 0 To 255
59950   lngC = rgb(intR, intG, intB)
59960   pb.Line (0, sngY)-(pb.width, sngY), lngC
59970   sngY = sngY + sngInc
59980 Next

59990 For intR = 255 To 0 Step -1
60000   lngC = rgb(intR, intG, intB)
60010   pb.Line (0, sngY)-(pb.width, sngY), lngC
60020   sngY = sngY + sngInc
60030 Next

60040 For intR = 0 To 255
60050   lngC = rgb(intR, intG, intB)
60060   pb.Line (0, sngY)-(pb.width, sngY), lngC
60070   sngY = sngY + sngInc
60080 Next

60090 For intG = 255 To 0 Step -1
60100   lngC = rgb(intR, intG, intB)
60110   pb.Line (0, sngY)-(pb.width, sngY), lngC
60120   sngY = sngY + sngInc
60130 Next

End Sub
