<div align="center">

## Releasing COM Objects from Memory


</div>

### Description

Releasing COM Objects from Memory.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[S Pandian](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/s-pandian.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB\.NET
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__10-7.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/s-pandian-releasing-com-objects-from-memory__10-5330/archive/master.zip)





### Source Code

<b><font face="Verdana" size="2"><font color="#0033CC">
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">Releasing COM Objects from Memory</font> :</font>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0"></b><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; We have created an instance of Excel in VB.NET has a habit of hanging around once we've finished with them. OK&nbsp;</font></p>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">But, Is there any big issue? , YES.&nbsp;</font></p>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"> &nbsp;&nbsp;&nbsp;
In VB6.0 , <font color="#0000FF"> Set Obj =
NOTHING</font>, Will release the memory. So why doesn't work for .NET ?.&nbsp;</font></p>
<p style="line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"> Referencing these objects within our .NET applications is that the actual COM code and its executable is outside of the managed. So, Garbage Collector will not be responsible for this kind of memory managements.&nbsp;</font></p>
<p><font face="Verdana" size="2" color="#0033CC"><b>Runtime Callable Wrapper :&nbsp;</b></font></p>
<p>&nbsp;&nbsp;&nbsp; <font face="Verdana" size="2">When reference the COM objects in our .NET, automatically wraps those references in something called RCW.
'UNMANAGED' objects dealt with within the 'MANAGED' .NET environment thru the <b>RCW</b>(<b>R</b>untime
<b>C</b>allable <b>W</b>rapper).&nbsp;</font></p>
<p>&nbsp;&nbsp;&nbsp;<font face="Verdana" size="2"> The only concession is that we must release each RCW as part of our cleanup process through the method named "<font color="#0000FF">Marshal.ReleaseComObject</font>" available in "<font color="#0000FF">System.Runtime.InteropServices</font>" namespace.&nbsp;</font></p>
<p>&nbsp;&nbsp;&nbsp;<font face="Verdana" size="2"> WE can rely on the garbage collector to automatically perform the necessary memory
management tasks. However, unmanaged resources require explicit cleanup.&nbsp;</font></p>
<p><font face="Verdana" size="2"><font color="#0033CC"><b>
Code Sample :</b></font>&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
Private objEx As Excel.Application&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
Private objWB As Excel.Workbook&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
Private objWS As Excel.Worksheet&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
objEx = New Excel.Application()&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
objWB = objEx.Workbooks.Add()&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
objWS = objWB.Worksheets.Add&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
objEx.Visible = True&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
Dim intRow As Integer&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
For intRow = 1 To 7&nbsp;</font>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
objWS.Range("A" & intRow).Value = Date.Today.AddDays(intRow).ToString("dddd")&nbsp;</font></p>
<p style="line-height: 100%; background-color: #EEEEEE; margin-top: -1; margin-bottom: 0"><font face="Verdana" size="2">
Next&nbsp;</font></p>
<p><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; It Release the Objects from the Memory , But not from the
<font color="#0000FF"> Task Manager</font>. <b>(</b><i>Each time we are closing the application then the following "ReleaseComObject" only enough other wise "Gc.Collect()" is also
needed.</i><b>)</b></font></p>
<p style="background-color: #EEEEEE; line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">System.Runtime.InteropServices.Marshal.<font color="#0000FF"><b>ReleaseComObject</b></font>(objEx)</font></p>
<p style="background-color: #EEEEEE; line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">System.Runtime.InteropServices.Marshal.<font color="#0000FF"><b>ReleaseComObject</b></font>(objWB)</font></p>
<p style="background-color: #EEEEEE; line-height: 100%; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">System.Runtime.InteropServices.Marshal.<font color="#0000FF"><b>ReleaseComObject</b></font>(objWS)</font></p>
<p><font face="Verdana" size="2">&nbsp;It Force to Release the Objects from the <font color="#0000FF"> Task Manager</font> Also.
<b> (</b> <i> It is best to use when the application is open, But we want to release the Object from Memory & Task
Manager</i><b>)</b>&nbsp;</font></p>
<p style="background-color: #EEEEEE"><font face="Verdana" size="2">Gc.<font color="#0000FF"><b>Collect</b></font>()&nbsp;</font></p>
<p><font face="Verdana" size="2"><b><font color="#FF0000">
Note : </font></b>Better to use this method to force the system to attempt to reclaim the maximum amount of available memory.</font></p>

