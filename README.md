<div align="center">

## How to Use DyCOMHelper to achieve Real inheritance in vb6\.\.etc \[UPDATE\]


</div>

### Description

You can download entire demo file from this url:http://www.ph586.net/soft/TestDyCOMHelperLib.zip

Before you use this demo, you should do by the following steps,

1.Include DyCOMHelperType.tlb and TestInheritLib.tlb under Types directory to this project.

2.You needn't to include DyCallerLib.dll and DyCOMHelperLib.dll.because they are standard dlls.

3.Start with pressing F8 and see how it works

Somthing improfect:

you need to create type Lib manually by the tool which can be found in the CD distributed with Advanced Visual Basic 6 written by Matthew Curland.

More information can be found on the website http://www.powervb.com/

we will provide tlb creation tools which can transfer VB6 Codes to type lib in the future .

There are three classes which are Animal, Wolf and Demiwolf. Demiwolf inheres from wolf, wolf inherets from animal and animal is derived from IDispatch.

This is to show you how to achieve real inherence in VB6, how to create object and use it. It will bring you to the world of COM at the back of VB6. Let us have a look at it.

Features of object created by DyCOMHelper:

1.Object is written in Moudle(.bas), not in Class Moudle(.cls).

2.Object is lightweight. Every instance ocupies 20 bytes in memory which is about 20% of VB6's(at least 96 bytes per instance).

3.Type is defined in TypeLib(.tlb file) currently.

4.It can be used like object of VB6. Even though it is not created by VB6 object system, it can be identified by VB6 as object.

5.It supports late binding, error handling, multible interface in one class, etc..

6.It supports real inherance which will be explained by this demo.

7.Speed of function invoking in IDE is 21% faster than that of VB6 Class. After you complied, it is 50% faster than that of VB6 Class.

8.Speed of instance creation is 10 times faster maximum.

9.Speed of instance release is about 100 times faster maximum.

10.Instance creation mode is richer than VB6. Instance of VB6 Class is created on heap whose spead is very low. While, DyComHeper could let you create your object on either heap or stack. DyComHeper can also support you to create from certain struct(UDT) or fix size memory manager supplied by DyComHelper which is written by Mathew Curland.

11.We supply you real pointer access to the inner data of an object created by DyComHelper with faster speed.

12.Creating thousands of classes by DyCOMHelper, you won't need to worry about efficiency.

You will find more good features about it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2011-01-08 17:19:16
**By**             |[Coobila](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/coobila.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[How\_to\_Use2195841112011\.zip](https://github.com/Planet-Source-Code/coobila-how-to-use-dycomhelper-to-achieve-real-inheritance-in-vb6-etc-update__1-73678/archive/master.zip)








