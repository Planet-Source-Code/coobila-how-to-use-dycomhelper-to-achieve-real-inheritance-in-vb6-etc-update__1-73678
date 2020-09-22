Attribute VB_Name = "modDemiwolf"
Option Explicit


Public Type dyDemiwolfType
  Master                 As String
  SheepCount             As Integer
End Type
  

'----------------------------------------------------------------------------------------------------
'DemiWolf
'----------------------------------------------------------------------------------------------------
    Private Sub pClass_Initialize_DemiWolf(ByRef This As dyDemiwolfType)
      g.Count = g.Count + 1
    End Sub

    Private Sub pClass_Terminate_DemiWolf(ByRef This As dyDemiwolfType)
      g.Count = g.Count - 1
      Dim Destroy        As dyDemiwolfType
      This = Destroy
    End Sub
    
    Private Function This_Prop_GetMaster(ByRef This As dyDemiwolfType, ByRef Ret As String) As HResult
      Ret = This.Master
    End Function
    
    Private Function This_Prop_LetMaster(ByRef This As dyDemiwolfType, ByRef Newval As String) As HResult
      This.Master = Newval
    End Function
    
    Private Function This_Prop_GetSheepCount(ByRef This As dyDemiwolfType, ByRef Ret As Integer) As HResult
      Ret = This.SheepCount
    End Function
    
    Private Function This_Prop_LetSheepCount(ByRef This As dyDemiwolfType, ByRef Newval As Integer) As HResult
      This.SheepCount = Newval
    End Function

    Private Function This_Func_ToString(ByRef This As dyDemiwolfType, ByRef Ret As String) As HResult
      Ret = "Demiwolf"
    End Function
    
    Public Sub zInitDemiWolf()
      If Types.Demiwolf Is Nothing Then
        Dim tThis                  As dyDemiwolfType
    
        Set Types.Demiwolf = g.Helper.NewCOMType(INTERFACES_2, LenB(tThis), Types.Wolf, dyIABAll, _
                                                dyIAOHeap Or dyIAOMemory, 8, AddressOf pClass_Initialize_DemiWolf, _
                                                AddressOf pClass_Terminate_DemiWolf, , , , VarPtrArray(g.Ptrs.Demiwolfs))
     
        Call Types.Demiwolf.ImplementsInterface(g.TypeLib, "DemiWolf", True, _
                                                 AddressOf This_Prop_GetMaster, _
                                                 AddressOf This_Prop_LetMaster, _
                                                 AddressOf This_Prop_GetSheepCount, _
                                                 AddressOf This_Prop_LetSheepCount)
                                                       
        Call Types.Demiwolf.ImplementsInterface(g.TypeLib, "IObject", True, _
                                                 AddressOf This_Func_ToString)
                                                       
      End If
    End Sub






