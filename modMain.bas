Attribute VB_Name = "modMain"
'DyCOMHelper--Real inheritance in vb6,create lightweight COM object in a simple way
'This codes is free for personal use.If you use it in your business project, please contact me by my email: justthisone@hotmail.com
'Before you use this demo, you should do by the following steps,
'1.Include DyCOMHelperType.tlb and TestInheritLib.tlb under Types directory to this project.
'2.You needn't to include DyCallerLib.dll and DyCOMHelperLib.dll.because they are standard dlls.
'3.Start with pressing F8 and see how it works
'
'Somthing improfect:
'  you need to create type Lib manually by the tool which can be found in the CD distributed with Advanced Visual Basic 6 written by Matthew Curland.
'  More information can be found on the website http://www.powervb.com/
'
'  we will provide tlb creation tools which can transfer VB6 Codes to type lib in the future .
'
'There are three classes which are Animal, Wolf and Demiwolf. Demiwolf inheres from wolf, wolf inherets from animal and animal is derived from IDispatch.
'
'This is to show you how to achieve real inherence in VB6, how to create object and use it. It will bring you to the world of COM at the back of VB6. Let us have a look at it.
'Features of object created by DyCOMHelper:
'1.Object is written in Moudle(.bas), not in Class Moudle(.cls).
'2.Object is lightweight. Every instance ocupies 20 bytes in memory which is about 20% of VB6's(at least 96 bytes per instance).
'3.Type is defined in TypeLib(.tlb file) currently.
'4.It can be used like object of VB6. Even though it is not created by VB6 object system, it can be identified by VB6 as object.
'5.It supports late binding, error handling, multible interface in one class, etc..
'6.It supports real inherance which will be explained by this demo.
'7.Speed of function invoking in IDE is 21% faster than that of VB6 Class. After you complied, it is 50% faster than that of VB6 Class.
'8.Speed of instance creation is 10 times faster maximum.
'9.Speed of instance release is about 100 times faster maximum.
'10.Instance creation mode is richer than VB6. Instance of VB6 Class is created on heap whose spead is very low. While, DyComHeper could let you create your object on either heap or stack. DyComHeper can also support you to create from certain struct(UDT) or fix size memory manager supplied by DyComHelper which is written by Mathew Curland.
'11.We supply you real pointer access to the inner data of an object created by DyComHelper with faster speed.
'12.Creating thousands of classes by DyCOMHelper, you won't need to worry about efficiency.
' You will find more good features about it.

'
'------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Public Declare Function GetCOMHelper Lib "DyCOMHelperLib.dll" (Optional ByVal lStdCallerPath As Long) As COMHelper
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Type dygTypesType
  Animal                  As DyComHelperType.COMType
  Wolf                    As DyComHelperType.COMType
  Demiwolf                As DyComHelperType.COMType
End Type
  'All types
  
Public Type dygPtrsType
  Animals()                As dyAnimalType
  Wolfs()                  As dyWolfType
  Demiwolfs()              As dyDemiwolfType
End Type

Public Type dygVariablesType
  TypeLib                 As DyComHelperType.TypeLib                    '//DyTestInheritType. The TypeLib is used in this demo. Animal\Wolf\DemiWolf was defined in this typelib.
  Helper                  As DyComHelperType.COMHelper
  Boost                   As DyComHelperType.Boost
  Ptrs                    As dygPtrsType
  Count                   As Long                                       '//Used to record the creation and releasing for object. If it is 0 at last ,it means that all objects are released correctly.
End Type
Public g                  As dygVariablesType
Public Types              As dygTypesType


Sub Main()
  Call pInitial
    '//To Use the helper ,we need to initial it before using.
    '//In this procedure,types info(Animal,Wolf,Demiwolf) would be initialized just after the Helper is initialized.
  
  
  Call pTestHelloWorld
    '//Just like all others foundation demos,we start this demo with Helloworld function
    
  Call pTestCreation
    '//Tell you how to create objects in tradtional way,pointer mode ,etc.
    
  Call pTestRelease
    '//Tell you how to release your objects created by DyCOMHelper
  
  Call pTestObjectFeature
    
  Call pTestStealData
    '//Get an instance ,use it,release it, but no QueryInterface\Addref\ReleaseRef will be occured
    
  Call pTestInherit
    '//Show you the feature of inheritance.
  
  Call pTestMultipleInterfaceSupport
    '//Show you how DyCOMHelper supports Multiple interfaces in one class
    
  Call pTestAccessInnerData
    '//Show you how to access the inner data of an object created by DyCOMHelper
    
  Call pTestClone
    '//Another demo to show you how to asscess inner data
    '//This example shows you how to clone an object.
    
  Call pTestConstructorWithArgs
    '//Show you how to add parameter to constructor function
    
  Call pTestAccessMeAndBase
    '//How to use Me like that in vb6.How to access base from the derived
    
  Call pTestAccessDerived
    '//How to access the derived from base
    
  Call pTestErrorCatch
    '//How to return an error from class.Please be noted of the difference between DyCOMHelper and VB6.
    
  Call pTestDispatch
    '//Testing of late bindding
    
  Call pTestShareData
    '//Try to share data between two objects
    
  Call pTestAllocOnStack
    '//Try to create an object from stack. Before you use this mode ,you need to make sure that you have enough control of the object.Otherwise, it will be very dangous.
    '//Do not set break pointer in this procedure.
                                            
  
Over:
  MsgBox g.Count
    '//If we get 0,it means that all objects are release correctly.
  
  'Call pTerminate
End Sub

Private Sub pTestObjectFeature()
  Dim oDemiwolf     As Demiwolf
  Dim oAnimal       As Animal
  Dim oObj          As Object
  Dim oIObj         As IObject
  

  Set oDemiwolf = Types.Demiwolf.NewInstance
    '//We create an object to test
  
  If TypeOf oDemiwolf Is IObject Then
    Debug.Print "Demiwolf is an IObject type object"
  End If
  
  Set oObj = oDemiwolf
    '//This line of code will run well because Demiwolf is derived from IDispatch which is Equal with Object in vb6.
    
  Set oAnimal = oDemiwolf
    '//This line of code will run well because a Demiwolf is an Animail
    
  Debug.Print oAnimal Is oDemiwolf
    '//It will output true. If we query IUnknown interface on oAnimal and on oDemiwolf,it will return the same pointer.
    '//So that we can say oAnimal Is oDemiwolf
  Debug.Print g.Helper.GetRefCount(oAnimal), g.Helper.GetRefCount(oDemiwolf)
    '//They both will output 3
    
  Debug.Print TypeOf oAnimal Is IObject
    '//It will output True.Because IObject interface is implemented on Animal Class.
    
  Set oDemiwolf = Nothing
  Set oObj = Nothing
    '//The object has not been released because one reference is held by oAnimal.
    
  Set oAnimal = Nothing
    '//The object will be released.
    
    
  Set oDemiwolf = Types.Demiwolf.NewInstance
  Set oIObj = oDemiwolf
  Debug.Print oIObj.ToString
    '//Just like vb6's class, this line will run well because IObject interface is implememted by Demiwolf class.
    '//But this line of code is not so efficient,because one time QueryInterface\ReleaseRef and 2 times AddRef will occur.
  Set oIObj = Nothing
  
  Call g.Boost.Assign(oIObj, g.Helper.GetInstanceReturnPtr(oDemiwolf, INTERFACE_1, 0))
  Debug.Print oIObj.ToString
    '//This line of code will get the same result,but it will be more efficient,because no QueryInterface\ReleaseRef\Addref will occur
    'Function GetInstanceReturnPtr(oDyInstance As Any       The object we want to get from. We can pass the object or the pointer of the object
    ' lInterfaceIndex As dyInterfaceIndexEnum               The interface index we want to get. In this example,the IObject interface is implemented on the Index of 1(Zero Base).
    '[cRefToAdd As Long]                                    The ref count you would like to add before this function returns.
  Call g.Boost.AssignZero(oIObj)
    '//'Set oIObj = Nothing' would not run well because there is no addref occured when we set the object pointer to the variable oIObj. But 'Set oIObj = Nothing' will occur ReleaseRef for one time which will break the balance of refcount.
   
End Sub

Private Sub pTestRelease()
  Dim oDemiwolf     As Demiwolf
  Dim pDemiwolf     As Long
  
TestTraditional:
  Set oDemiwolf = Types.Demiwolf.NewInstance
  Set oDemiwolf = Nothing
    '//Release the object just created by traditional way.We suggest you use this way.The object life cycle will be under control just like VB's object.
        
OthersWay:
  pDemiwolf = Types.Demiwolf.NewInstanceReturnPtr
    '//An object is created and one reference will be held by this object
  Call g.Boost.Assign(oDemiwolf, pDemiwolf)
  oDemiwolf.Color = vbWhite:  Debug.Print oDemiwolf.Color
  Call g.Boost.AssignZero(oDemiwolf)
    '//The variable oDemiwolf will be destroyed
    '//But the object is not released because one reference is still held by the object
    
  Call g.Helper.ReleaseRef(pDemiwolf)
    '//This step will invoke the modDemiwolf.pClass_Terminate_DemiWolf function and the object will be released.
  
End Sub

Private Sub pTestCreation()
  Dim oDemiwolf     As Demiwolf
  Dim oDemiwolf2    As Demiwolf
  Dim oWolf         As Wolf
  Dim oAnimal       As Animal
  Dim pDemiwolf     As UnkHandle            '//UnkHandle,an alias to long
  Dim oObj          As IObject

TestTraditional:
  Set oDemiwolf = Types.Demiwolf.NewInstance
  Set oDemiwolf = Nothing
    '//Traditional ways.
      '//There are only 2 lines as you can see,but it will lead to QueryInterface  at least once and AddRef/ReleaseRef some times
      '//The key word New achieved at the back of VB6 is like this generally.
      
TestTraditional2:
  Call Types.Demiwolf.CreateInstance(oDemiwolf)
  Set oDemiwolf = Nothing
    '//In this way it will be faster than the last way. It will save one temp variable and avoid AddRef/ReleaseRef for one time.

TestPointerMode:
  pDemiwolf = Types.Demiwolf.NewInstanceReturnPtr(1, INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT)
    '//Create an object and return the pointer.
    'Function NewInstanceReturnPtr(
    '  [cRef As Long = 1]                                   the reference count of the new object.it will be 1 in default.
    '  [lInheritLevelRet As dyInheritLevelOffsetEnum]       which level of object in the inheritance index you would like to return. It will return the current type(Demiwolf ) if we ignore this parameter.
    '                                                       if we pass INHERIT_LEVEL_OFFSET_1 ,it will return the pointer of the Wolf Object.
    '                                                       if we pass INHERIT_LEVEL_OFFSET_2,it will return the pointer of the Animal Object.
    '  [lInterfaceIndexRet As dyInterfaceIndexEnum]         which interface you would like to return.
    '                                                       There are 2 interface is implemented in every class in this demo.
    '                                                       For example,Animal,we implement Animal interface in default,and IObject is implemented additionally.
    '                                                       pDemiwolf = Types.Demiwolf.NewInstanceReturnPtr(1, INHERIT_LEVEL_OFFSET_0_None, INTERFACE_1)
    '                                                       An Object whose type is IObject will be returned.
    '  [lThisStruct As Long]                                This is for advanced user.if a COMType can support creating object and it's inner data on different memory,not continual Memory,it will be very useful.
    '                                                       I will provide some examples in the next demo
    '  [lAddressOfDefaultConstructorOfBase As Long]         we can define the Base Obejct constructor function you would like to construct the inner base object.
    '  [oBaseInstance As Any])                              we can provide the base object.you can pass the base object or it's pointer.
    '                                                       if we ignore this parameter,an inner base object will be created by COMHelper automatically
    '                                                       Example:
    '                                                           Set oWolf = Types.Wolf.NewInstance
    '                                                           Set oDemiwolf = Types.Wolf.NewInstance(, , , , oWolf)
  Call g.Boost.Assign(oDemiwolf, pDemiwolf)
    '//Set the pointer value to a variable whose type is Demiwolf.
    '//Boost.Assign will set the value to the variable directly without add/release ref or QueryInterface,so that, we need to confirm that the pDest is nothing.
    


    '  We can write the codes in the simple way just like this:
    '  Call g.Boost.Assign(oDemiwolf, Types.Demiwolf.NewInstanceReturnPtr(1,INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT))

    '  Sum-Up:
    '        1. in the whole creation process,there are no AddRef,ReleaseRef or QueryInterface was occured.All system resources are only used on object creation.
    '           Please be noted that one QueryInterface will definitely occur both AddRef and ReleaseRef one time.Otherwise, it will make a series of UUID comparason.
    '           This is a nightmare to the classes which support multiple interfaces and has inheritance behavior.
    '           The reason of slow initialize of UserControl is because most time is used to QueryInterface which is meaningless but necessary .
    '           Usercontrol implements too many interfaces which are hidden by VB
    '        2.If we use pointer mode just like the above code,we can access the inner data by the pointer,pass the object by the pointer without worrying redundant invoke of AddRef,ReleaseRef or QueryInterface.

TestAssignAddref:
    Call g.Boost.AssignAddRef(oDemiwolf2, oDemiwolf)
      '//It is equal with the keyword Set ,it will only occur Addref,but no QueryInterface.
      '//Please be noted that the type of pSrc must be equal to pDest.
    Set oDemiwolf2 = Nothing
      '//Release it
End Sub

Private Sub pTestErrorCatch()
  Dim oAnimal           As Animal
  
  Call g.Boost.Assign(oAnimal, Types.Animal.NewInstanceReturnPtr(1, INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT))
  On Error Resume Next
  oAnimal.LegsCount = 5
    '//Please be noted of the code in modAnimal.This_Prop_LetLegsCount
    
  Debug.Print Err.Description
  
End Sub

Private Sub pTestStealData()
  Dim oDemiwolf     As Demiwolf
  Dim oDemiwolf2    As Demiwolf
  
  '----------------------------------------------------------------------------------------------------------------
  ' This is only for demo.I belive nobody will write code like this.
  '
  '----------------------------------------------------------------------------------------------------------------
  Call g.Boost.Assign(oDemiwolf, Types.Demiwolf.NewInstanceReturnPtr(1, INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT))
    
    
TestTraditional:
  Set oDemiwolf2 = oDemiwolf
    '//It will evoke Addref 1 time.
  oDemiwolf2.LegsCount = 4
  Set oDemiwolf2 = Nothing
    '//It will evoke ReleaseRef 1 time.
    
TestDyCOMHelperMode:
  Call g.Boost.Assign(oDemiwolf2, oDemiwolf)
     '//No Addref
  oDemiwolf2.LegsCount = 4
    '//Steal it and use
  Call g.Boost.AssignZero(oDemiwolf2)
    '//Realse it ,No RealeaseRef
    '------------------------------------------------------------------------------------------------------------------------------------
    ' Note:since reference count is not added,if the object is released during the using of oDemiwolf2, the system will crash.
    '
    '------------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub pTestAccessDerived()
  Dim oDemiwolf     As Demiwolf
  Dim oWolf         As Wolf
  Dim oAnimal       As Animal
      
  Set oDemiwolf = Types.Demiwolf.NewInstance
  Set oWolf = oDemiwolf
  Set oAnimal = oDemiwolf
  
  Debug.Print g.Helper.GetDerivedPtr(oWolf), g.Boost.ObjptrForUnknown(oDemiwolf), ObjPtr(oDemiwolf)
    '//This will output the same value.
    
  Debug.Print g.Helper.GetDerivedPtr(oAnimal, , INHERIT_LEVEL_OFFSET_1), g.Boost.ObjptrForUnknown(oDemiwolf), ObjPtr(oDemiwolf)
    '//This will output the same value.
End Sub

Private Sub pTestAccessBase()
  Dim oDemiwolf     As Demiwolf
  Dim oWolf         As Wolf
  Dim oAnimal       As Animal
      
  Set oDemiwolf = Types.Demiwolf.NewInstance
  Set oWolf = oDemiwolf
  Set oAnimal = oDemiwolf
  
  Debug.Print g.Helper.GetBasePtr(oDemiwolf), g.Boost.ObjptrForUnknown(oWolf)
    '//This will output the same value.
    
  Debug.Print g.Helper.GetBasePtr(oDemiwolf, , INHERIT_LEVEL_OFFSET_1), g.Boost.ObjptrForUnknown(oAnimal)
    '//This will output the same value.
  

  Set oWolf = Nothing
  Set oAnimal = Nothing
  
  Call g.Boost.Assign(oWolf, g.Helper.GetBasePtr(oDemiwolf))
  oWolf.Color = vbWhite
  Call g.Boost.AssignZero(oWolf)
    '//Show you some skills to use pointer
  
End Sub

Private Sub pTestHelloWorld()
  Dim oHelloWorld       As Demiwolf
  
  Set oHelloWorld = Types.Demiwolf.NewInstance
    '//Create an object whose type is Demiwolf
  
  oHelloWorld.Color = vbWhite
  Debug.Print oHelloWorld.Color
    '//Access members of the object
  
  Set oHelloWorld = Nothing
    '//Release it
End Sub


Private Sub pTestInherit()
  Dim oDemiwolf     As Demiwolf
  Dim oWolf         As Wolf
  Dim oAnimal       As Animal
    
  Call g.Boost.Assign(oDemiwolf, Types.Demiwolf.NewInstanceReturnPtr())
    
  oDemiwolf.LegsCount = 4
  Debug.Print oDemiwolf.LegsCount
    '//LegsCount is a member of Animal,this line of code will jump to modAnimal.This_Prop_GetLegsCount
    
  oDemiwolf.Color = vbWhite
  Debug.Print oDemiwolf.Color
    '//Color is a member of Wolf ,this line of code will jump to modWolf.ThisProp_GetColor
   
  oDemiwolf.Master = "DyCOMHelperLib"
  Debug.Print oDemiwolf.Master
    '//Master is a member of Demiwolf ,this line of code will jump to modDemiwolf.This_Prop_GetMaster
   
  Set oWolf = oDemiwolf
    '//This line of code will run well because a Demiwolf is Derived from Wolf,so we can say that a Demiwolf is a Wolf
  Debug.Print oWolf.LegsCount
End Sub

Private Sub pTestClone()
  Dim oAnimal           As Animal
  Dim oClone            As Animal
  Dim oArrayOwner       As ArrayOwner
  Dim tAnimal()         As dyAnimalType
  
  
  Call g.Boost.Assign(oAnimal, NewAnimalReturnPtr(4, 2))
    '//Create an object to test
  Set oArrayOwner = g.Helper.NewArrayOwner(VarPtrArray(tAnimal), Types.Animal.LenBOfThisStruct, 1)
    '//Create an arrayowner to access the inner data
    
  Call oArrayOwner.Attach(Types.Animal.ThisStructPtrFromMe(oAnimal))
    '//Types.Animal.ThisStructPtrFromMe(oAnimal) will return the pointer of the inner data of oAnimal
    '//oArrayOwner.Attach will Map the inner data to the Array tAnimal
  
  Call Types.Animal.CreateInstance(oClone)
  Call Types.Animal.AttachThisStructTo(oClone)
  g.Ptrs.Animals(0) = tAnimal(0)
    '//Set value to the inner data
  
  Debug.Print oClone.EyesCount, oClone.LegsCount
  
  Set oArrayOwner = Nothing
    '//We should release the oArrayowner
End Sub

Private Sub pTestAccessInnerData()
  Dim oAnimal           As Animal
  
  Call g.Boost.Assign(oAnimal, modAnimal.NewAnimalReturnPtr(4, 2))
    '//We create an Animal object with 4 legs and 2 eyes
    
  Call Types.Animal.AttachThisStructTo(oAnimal)
    '//Attach the inner data of oAnimal to the array g.Ptrs.Animals
  
  'With g.Ptrs.Animals(0)
  g.Ptrs.Animals(0).EyesCount = g.Ptrs.Animals(0).EyesCount + 2
  g.Ptrs.Animals(0).LegsCount = g.Ptrs.Animals(0).LegsCount + 4
  'End With
    '//Access the inner data in the way you like.
    '//We do not suggest using keyword With because the keyword With will lock the Array
  
  With oAnimal
    Debug.Print .EyesCount
    Debug.Print .LegsCount
  End With
    '//Out put the data we just modified
End Sub

Private Sub pTestConstructorWithArgs()
  Dim oAnimal           As Animal
  
  Call g.Boost.Assign(oAnimal, modAnimal.NewAnimalReturnPtr(4, 2))
  With oAnimal
    Debug.Print .EyesCount
    Debug.Print .LegsCount
  End With
End Sub

Private Sub pTestDispatch()
  Dim oWolf          As Wolf
  Dim oObj           As Object

  Set oWolf = Types.Wolf.NewInstance
  oWolf.EyesCount = 4
    '//VTable binding mode ,Not late binding.

  Set oObj = oWolf
  oObj.LegsCount = 4
  oObj.EyesCount = 2
    '//Late binding mode
  
  'On Error Resume Next
  Call CallByName(oObj, "EyesCount", VbLet, 2)
  Debug.Print CallByName(oObj, "EyesCount", VbGet)
  Debug.Print CallByName(oObj, "EyesCount", VbLet, 2)
    '//Use the Callbyanme
  
End Sub

Private Sub pTestMultipleInterfaceSupport()
  Dim oWolf          As Wolf
  Dim oObj           As IObject
  
  Set oWolf = Types.Wolf.NewInstance
  oWolf.EyesCount = 4
  oWolf.Origin = "China"
 
  Set oObj = oWolf
  Debug.Print oObj.ToString
    '//This line of code will run well
    
End Sub

Private Sub pTestAccessMeAndBase()
  Dim oWolf          As Wolf
  
  Set oWolf = Types.Wolf.NewInstance
  
  oWolf.Color = vbWhite
    '//In the implement of Wolf.Color(Get) will show you how to get Base instance.
    
  Debug.Print oWolf.Color
    '//In the implement of Wolf.Color(Let) will show you how to get Me instance.
End Sub

Private Sub pTestShareData()
  Dim a            As Demiwolf
  Dim b            As Demiwolf
  Dim tData        As dyDemiwolfType
  
  Set a = Types.Demiwolf.NewInstance(INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT, VarPtr(tData))
  Set b = Types.Demiwolf.NewInstance(INHERIT_LEVEL_OFFSET_0_None, INTERFACE_0_DEFAULT, VarPtr(tData))
    '//a and b are both created on the same inner data.
    '//so that we can share data between different objects.
    
  a.Master = "Master Shared"
  
  Debug.Print b.Master, tData.Master
End Sub


Private Sub pTestAllocOnStack()
  Dim a            As Demiwolf
  Dim lPtr         As Long
  
  '-----------------------------------------------------------------
  '  *****  Do not set break pointer in this procedure.
  '-----------------------------------------------------------------
  
  lPtr = g.Boost.StackAllocZero(Types.Demiwolf.LenBPerInstance)
    '//This will alloc Types.Demiwolf.LenBPerInstance size of memory on stack
      '//Do not use g.Boost.StackAlloc,because we should zero the memory before we use it
      '//Please be noted that we should Alloc just after the variables declaration.
  
  Set a = Types.Demiwolf.NewInstanceFromMemory(lPtr)
    '//We create an object use the memory alloced before
    '//The variable a can be passed to any function.However, one point you need to make sure that the refcount will be the same with what it is before we passed it out.
  a.Master = "The Master"
    
  Debug.Print a.Master
  Set a = Nothing
    '//Before we free the memory,we should release the object.
  
  Call g.Boost.StackFree(Types.Demiwolf.LenBPerInstance)
    '//Free the memory we alloced on stack.

End Sub

Private Sub pInitial()
  Dim sAppPath   As String
  '
  sAppPath = App.Path: If VBA.Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
  
  Set g.Helper = GetCOMHelper(VarPtr(sAppPath & "DyCallerLib.dll"))
    '//Initial the helper
    '//Note:we need to use varptr to pass a String parameter to the GetCOMHelper function.
    
  Set g.Boost = g.Helper.Boost
    '//Get the Boost object which is used in this demo.
  
  Set g.TypeLib = g.Helper.NewTypeLib(sAppPath & "Types\TestInheritLib.tlb")
  Debug.Print g.TypeLib.Guid, g.TypeLib.Name, g.TypeLib.lcid, g.TypeLib.MajorVerNum, g.TypeLib.MinorVerNum, g.TypeLib.OsKind = dyOSWIN32
    '//Initial the Type lib which is used in this demo.  Animail\Wolf\DemiWolf is defined in this type lib.
  
  Call modAnimal.zInitAnimal                '//Please see the achivement in modAnimal.zInitAnimal, which will show you how to ceate a COMType,and implement interface for it.
  Call modWolf.zInitWolf
  Call modDemiwolf.zInitDemiWolf
    '//Before you create instance and use them you should initial COMTypes At first.
    
End Sub

Private Sub pTerminate()
  Set Types.Demiwolf = Nothing
  Set Types.Wolf = Nothing
  Set Types.Demiwolf = Nothing
  
  Set g.Boost = Nothing
  Set g.TypeLib = Nothing
  Call g.Helper.Unload
  Set g.Helper = Nothing
End Sub
