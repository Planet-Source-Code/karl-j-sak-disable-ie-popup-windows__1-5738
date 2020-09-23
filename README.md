<div align="center">

## Disable IE Popup Windows


</div>

### Description

Disables IE popup windows by using an VB Front-end browser and redirecting the popup window to a second invisible webbrowser control.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Karl J\. Sak](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karl-j-sak.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/karl-j-sak-disable-ie-popup-windows__1-5738/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
' navigate to a website, I suggest www.aol.com
WebBrowser1.Navigate "http://www.aol.com"
End sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
'this sets the popup window to another browser control
'in which webbrowser2.visible = false
Set ppDisp = WebBrowser2.Object
End Sub
```

