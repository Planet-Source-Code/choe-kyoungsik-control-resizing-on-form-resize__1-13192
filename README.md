<div align="center">

## Control Resizing on Form\_Resize


</div>

### Description

You know... Real resize function needs some information about what to do. It cannot be done without information. The program even doesn't know what I want until I give it!!!

For example, Let's say you have 4 controls.

First one is command box, so it should be moved on form_resize.

Second one is text box, so it should be wider or vice versa on form_resize.

Third one is multi-line text box, so it should vary on width and height both.

Fourth one is a label below the third one, so it should move vertically.

You can DEFINE very easily how your controls should resized according to the parent form's resizing.

If you are doing a lot of forms(maybe over 30, 50, 100?), this function will be very helpful.
 
### More Info
 
ListBox and DBList will be shrinking more and more when you maximize and minimize.

You should code additionally on Form_Resize,

like this:

If Me.Height > 1000 Then

ListBox1.Height = Me.Height - 300

End If

If one control has two child controls, its resize definition will be inherited from the first child on EvtFormResize() arguments.

You'll see what I mean if you take care of the frame control on the sample project.


<span>             |<span>
---                |---
**Submitted On**   |2000-11-28 09:47:42
**By**             |[CHOE KyoungSik](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/choe-kyoungsik.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1216711282000\.zip](https://github.com/Planet-Source-Code/choe-kyoungsik-control-resizing-on-form-resize__1-13192/archive/master.zip)








