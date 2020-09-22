Attribute VB_Name = "modAnimal"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ËµÃ÷£º
'     1¡¢¹ØÓÚÊôÐÔ¡¢·½·¨¡¢º¯ÊýµÄÉùÃ÷£º
'        Ã»ÓÐ½Ó´¥¹ýCOM±à³ÌµÄÈË£¬¿´µ½ Property Get LegsCount ±»ÉùÃ÷³Éº¯Êý£¬¿ÉÄÜ»á¾õµÃºÜÆæ¹Ö¡£
'        ÆäÊµ£¬ÕâÊÇÄúÏ°¹ßÁËVBÎªÎÒÃÇ·â×°ºÃµÄ´úÂë¡£ÆäÊµÔÚCOM²ãÃæÏÂ£¬Ëü×îÔ­Ê¼µÄÇéÐÎ¾ÍÊÇÕâÑùµÄ¡£
'        ±ê×¼µÄCOM£¬Ö»ÓÐº¯Êý£¨Sub ÆäÊµÒ²ÊÇº¯Êý£¬Ö»ÊÇ·µ»ØÁËVoid£©£¬Ã»ÓÐÊôÐÔºÍ·½·¨£¬º¯ÊýµÄ·µ»ØÖµ¿ÉÒÔÊÇÈÎºÎÀàÐÍ¡£
'        ¶øVB¶ÔÀà³ÉÔ±ÊµÏÖ£¬ËùÓÐ³ÉÔ±·µ»ØÖµÀàÐÍ¶¼ÊÇHResult£¬ÕâÀïÎÒÃÇÓÃ Long±íÊ¾£¬¸Ã·µ»ØÖµÓÃ×÷´íÎó´¦ÀíÓÃ¡£
'        Ô­ÀíÉÏ£¬Äú¿ÉÒÔÓÃDyCOMHelper´´½¨ÈÎÒâ·µ»ØÖµÀàÐÍµÄ£¬µ«ÊÇ£¬ÎÒÃÇ½¨ÒéÄú°´ÕÕVBµÄ¹æ·¶À´½øÐÐ¡£
'        ·½·¨µÄÉùÃ÷£º
'        Private Function This_Sub_SubName(ByRef This As YourStruct,Args......) As HResult
'        End Function
'        º¯ÊýµÄÉùÃ÷£º
'        Private Function This_Sub_FuncName(ByRef This As YourStruct,Args......,Ret As Type) As HResult
'        End Function
'        ÊôÐÔGetµÄÉùÃ÷£º
'        Private Function This_PropGet_PropName(ByRef This As YourStruct,Args......,Ret As Type) As HResult
'        End Function
'        ÊôÐÔLetµÄÉùÃ÷£º
'        Private Function This_PropLet_PropName(ByRef This As YourStruct,Args......,Newval As Type) As HResult
'        End Function
'        ÊôÐÔSetµÄÉùÃ÷£º
'        Private Function This_PropSet_PropName(ByRef This As YourStruct,Args......,Newval As Type) As HResult
'        End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Type dyAnimalType
  LegsCount              As Long
  EyesCount              As Long
End Type

Private mMemMgr          As FixedSizeMemoryManager

'----------------------------------------------------------------------------------------------------
'Animal
'----------------------------------------------------------------------------------------------------
    Public Function NewAnimalReturnPtr(ByRef lLegsCount As Long, ByRef lEyesCount As Long) As Long
       NewAnimalReturnPtr = Types.Animal.NewInstanceReturnPtr
       
       Call Types.Animal.AttachThisStructTo(NewAnimalReturnPtr)
       'With g.Ptrs.Animals(0)
       g.Ptrs.Animals(0).LegsCount = lLegsCount
       g.Ptrs.Animals(0).EyesCount = lEyesCount
       'End With
    End Function
    
    
    
    Private Function pClass_DefaultConstructor() As Long
      '//DefaultConstructor
       pClass_DefaultConstructor = Types.Animal.NewInstanceReturnPtr
    End Function
    
    Private Sub pClass_Initialize(ByRef This As dyAnimalType)
      g.Count = g.Count + 1
    End Sub

    Private Sub pClass_Terminate(ByRef This As dyAnimalType)
      g.Count = g.Count - 1
    End Sub
    
    Private Sub pMemoryFree(ByRef lPtr As Long)

    End Sub
    
    Private Function This_Prop_GetLegsCount(ByRef This As dyAnimalType, ByRef Ret As Long) As HResult
      Ret = This.LegsCount
    End Function
    
    Private Function This_Prop_LetLegsCount(ByRef This As dyAnimalType, ByRef Newval As Long) As HResult
      Select Case Newval
        Case 2, 4
          This.LegsCount = Newval
        Case Else
          This_Prop_LetLegsCount = g.Helper.Err.Raise(1001, "Animal.LegsCount(Let)", "Some error")
            '//Please be noted of the difference to Err.Raise in VB6
      End Select
    End Function
    
    Private Function This_Prop_GetEyesCount(ByRef This As dyAnimalType, ByRef Ret As Long) As HResult
      Ret = This.EyesCount
    End Function
    
    Private Function This_Prop_LetEyesCount(ByRef This As dyAnimalType, ByRef Newval As Long) As HResult
      'This.EyesCount = Newval
    End Function
    
    Private Function This_Sub_Move(ByRef This As dyAnimalType, ByRef lDistance As Double) As HResult
      Debug.Print "Animal_Sub_Move", lDistance
    End Function
    
    Private Function This_Sub_Bite(ByRef This As dyAnimalType) As HResult
      Debug.Print "This_Sub_Bite"
    End Function
    
    Private Function This_Func_ToString(ByRef This As dyDemiwolfType, ByRef Ret As String) As HResult
      Ret = "Animal"
    End Function
    
    Public Sub zInitAnimal()
      If Types.Animal Is Nothing Then
        Dim tThis                  As dyAnimalType
    
        Set Types.Animal = g.Helper.NewCOMType(INTERFACES_2, _
                               LenB(tThis), Nothing, dyIABAll, _
                               dyIAOHeap Or dyIAOMemory Or dyIAODefaultFixMemMgr, _
                               8, AddressOf pClass_Initialize, AddressOf pClass_Terminate, , , , _
                               VarPtrArray(g.Ptrs.Animals))
           ' Create the Animal type.
           'lInterfacesCount As Long                                        Number of interfaces
           'lLenBOfThisStruct As Long                                       Every object contains two parts:Object and it's inner data.In each implement function of member in this class, the first parameter 'This' will store the inner data
           'oInheritFrom As dyCOMHelperType.COMType                         COMType inherited from¡£
           'lInstanceAllocBehavior As dyInstanceAllocBehaviorEnum           Every object contains two parts:Object and it's inner data.This parameter will determine the memory alloc behavior for the two parts.
           '                                                                Both are created on continuous memory(dyIABAll) or create object and inner data separately.
           '                                                                Use this feature,you can share data between objects
           'lInstanceAllocOnSupport As dyInstanceAllocOnEnum                the memory alloc mode which will be supported by this ComType
           '                                                                dyIAOHeap                   the new object can be created on heap.for example: Set oNewAnimal = Types.Animal.NewInstance or Call g.Boost.Assign(oNewAnimal, Types.Animal.NewInstanceReturnPtr)
           '                                                                dyIAODefaultFixMemMgr       the new object can be created on the inner fix memory manager built in this COMType. For example:Set oNewAnimal = Types.Animal.NewInstanceFromDefaultFixMemMgr  or Call g.Boost.Assign(oNewAnimal, Types.Animal.NewInstanceFromDefaultFixMemMgrReturnPtr )
           '                                                                dyIAOFixMemMgr              the new object can be created on fix memory manager provided by yourself . lThisMemoryFreeFuncAddress must be appointed by yourself.Before the object release ,memory free procedure will be invoked.
           '                                                                dyIAOMemory                 the new object can be created on the memory appointed by yourself.
           'lInstancesPerBlock As Long                                      Number of instance  per block on default fix memory mgr.
           'lThisInitializeFuncAddress As Long
           'lThisTerminateFuncAddress As Long
           'lThisMemoryFreeFuncAddress As Long
           'lThisDefaultConstructFuncAddress As Long
           'lThisDefaultdDisposeFuncAddress As Long
           'pArrForPointerToAccessThisStruct                                If we appointe this parameter,we can use Types.Animal.AttachThisStructTo( oAnimal) to Attach the innerdata of  oAnimal,and access it by the array g.Ptrs.Animals
            

        Call Types.Animal.ImplementsInterface(g.TypeLib, _
                                              "Animal", True, _
                                              AddressOf This_Prop_GetLegsCount, _
                                              AddressOf This_Prop_LetLegsCount, _
                                              AddressOf This_Prop_GetEyesCount, _
                                              AddressOf This_Prop_LetEyesCount, _
                                              AddressOf This_Sub_Move, _
                                              AddressOf This_Sub_Bite)
         ' Implement the first interface.
         
        Call Types.Animal.ImplementsInterface(g.TypeLib, _
                                              "IObject", True, _
                                              AddressOf This_Func_ToString)
                                              
          '//Implement the second interface
      End If
    End Sub

