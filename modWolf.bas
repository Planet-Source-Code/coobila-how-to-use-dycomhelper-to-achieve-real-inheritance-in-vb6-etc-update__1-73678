Attribute VB_Name = "modWolf"
Option Explicit
Public Type dyWolfType
  Origin                 As String
  Color                  As OLE_COLOR
End Type


'----------------------------------------------------------------------------------------------------
'Wolf
'----------------------------------------------------------------------------------------------------
    Private Function DefaultConstructor() As Long
       DefaultConstructor = Types.Wolf.NewInstanceReturnPtr
    End Function
    
    Private Sub pClass_Initialize_Wolf(ByRef This As dyWolfType)
      g.Count = g.Count + 1
    End Sub

    Private Sub pClass_Terminate_Wolf(ByRef This As dyWolfType)
      g.Count = g.Count - 1
      Dim Destroy     As dyWolfType
      This = Destroy
    End Sub
    
    Private Function ThisProp_GetColor(ByRef This As dyWolfType, ByRef Ret As OLE_COLOR) As HResult
      Ret = This.Color
      
      Dim oBase           As Animal
        
      Call Types.Wolf.GetBase(oBase)        '
      Debug.Print oBase.LegsCount
        '//Get The Base
    End Function
    
    Private Function ThisProp_LetColor(ByRef This As dyWolfType, ByRef Newval As OLE_COLOR) As HResult
      This.Color = Newval
      
      Dim oMe              As Wolf
      
      Call Types.Wolf.GetMe(oMe)
      Debug.Print oMe.LegsCount
        '//Get Me(This)
      
    End Function
    
    Private Function ThisProp_GetOrigin(ByRef This As dyWolfType, ByRef Ret As String) As HResult
      Ret = This.Origin
    End Function
    
    Private Function ThisProp_LetOrigin(ByRef This As dyWolfType, ByRef Newval As String) As HResult
      This.Origin = Newval
    End Function
    
    Private Function This_Func_ToString(ByRef This As dyDemiwolfType, ByRef Ret As String) As HResult
      Ret = "Wolf"
    End Function
    

    Public Sub zInitWolf()
      If Types.Wolf Is Nothing Then
        Dim tThis                  As dyWolfType
    
        Set Types.Wolf = g.Helper.NewCOMType(INTERFACES_2, LenB(tThis), Types.Animal, _
                                             dyIABAll, dyIAOHeap, 8, _
                                             AddressOf pClass_Initialize_Wolf, AddressOf pClass_Terminate_Wolf, , _
                                             AddressOf DefaultConstructor, , VarPtrArray(g.Ptrs.Wolfs))
           '//Create Wolf type ,it inherits from the type Animal
      
        Call Types.Wolf.ImplementsInterface(g.TypeLib, "Wolf", True, _
                                            AddressOf ThisProp_GetColor, _
                                            AddressOf ThisProp_LetColor, _
                                            AddressOf ThisProp_GetOrigin, _
                                            AddressOf ThisProp_LetOrigin)
                                            
                                            
         
        Call Types.Wolf.ImplementsInterface(g.TypeLib, _
                                              "IObject", True, _
                                              AddressOf This_Func_ToString)

          
      End If
    End Sub

