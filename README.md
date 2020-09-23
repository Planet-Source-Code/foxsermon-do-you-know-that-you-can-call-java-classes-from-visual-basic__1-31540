<div align="center">

## Do you know that you can call Java classes from Visual Basic ??


</div>

### Description

I updated my article, cause for some reasons, Microsoft removed the SDK from its site. So finally I found it in another place (its include below) I did test it and worked it again. So if anyone wants to know how it works, just try it.

Any problem, dont hesitate to contact me.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[foxsermon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/foxsermon.md)
**Level**          |Intermediate
**User Rating**    |4.6 (46 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/foxsermon-do-you-know-that-you-can-call-java-classes-from-visual-basic__1-31540/archive/master.zip)





### Source Code

Do you know that you can call Java classes from Visual Basic ??<Br><Br>
let me guide you through this.
Firts at all, we need a Java class, <Br>
Here you will find an example:<br><br>
Java Code <br><Br>
//**************************************************
public class MyTest {<Br>
  public int myfunction(int value1, int value2) {<Br>
  return value1+value2;<br>
  }<Br>
}<Br>
//***************************************************
<Br><Br>
Compile it<br>
i.e. javac MyTest.java<br>
and when you get MyTest.class file, you must register it for this you will need Microsfot SDK for Java
you can download from this link <br>
http://www.microsoft.com/java/download/dl_sdk40.htm<Br>or<br> http://sf.gds.tuwien.ac.at/subcat/pro/webunitproj/<br>
when you have downloaded and installed. <Br>
include it on Path variable enviroment<Br>
just do it using Command Ms-DOS<Br><Br>
set path=%path%;C:\Program Files\Microsoft SDK for Java 4.0\bin\<br><br>
(this path could change)<Br><Br>
then you must register our MyTest.class file.<Br>
you should do it using javareg.exe file from Microsoft SDK for Java<Br><br>
i.e. javareg /register /class:MyTest /progid:MyTest<Br><br>
if everything is well-done then you should see a MessageBox displayed with <Br>
Succesfull register Class message.<Br>
Otherwise, you must check the correct spelling on the command line.<Br>
Now, the next step, copy the new Java class file that was generated.<Br>
MyTest.class<Br>
and paste in C:\Winnt\Java\Trustlib\ folder<Br>
if you have windows 98, this folder may change.<Br>
And the last step, open a New Project on Visual Basic.<Br>
And paste this brief code on the Form Load event for example and run it.<Br><Br>
 Set x = CreateObject("MyTest")<Br>
 MsgBox x.myfunction(1, 1)<Br><Br>
Congratulations !!! <Br><Br>
you can use Java Class from Visual Basic. <Br>
Please do not forget to vote for this article. =)<Br>

