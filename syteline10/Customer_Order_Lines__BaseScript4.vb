'//Form script from CustomerOrderLines [SYMIX_DEFAULT / [NULL]] - ancestor level: -04.

Option Explicit On
Option Strict On

Imports System
Imports Microsoft.VisualBasic
Imports Mongoose.IDO.Protocol
Imports Mongoose.Scripting
Imports System.Math

Namespace SyteLine.FormScripts
    Public Class CustomerOrderLines
        Inherits FormScript

#Region "MultiSiteItemSourcing"
        Sub StdFormCalledFormReturned()
            If (ThisForm.LastModalChildName = "MultiSiteItemSourcing") Then
                Dim SiteRef As String = ThisForm.Variables("vRecommendedSite").Value
                Dim Whse As String = ThisForm.Variables("vRecommendedWhse").Value

                If (Not (String.IsNullOrEmpty(SiteRef) Or String.IsNullOrEmpty(Whse))) Then
                    If (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite").ToString() <> SiteRef Or
                       ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Whse").ToString() <> Whse) Then
                        ThisForm.GenerateEvent("SetValueFromMultiSiteItemSourcing")
                    End If

                End If
            End If
        End Sub

#End Region



        Sub FindNumberOfModifiedRecords()
            Dim Ctr As Integer
            Dim NoOfRecords As Integer
            If ThisForm.PrimaryIDOCollection.IsCollectionModified Then
                For Ctr = 0 To ThisForm.PrimaryIDOCollection.GetNumEntries - 2
                    If ThisForm.PrimaryIDOCollection.IsObjectModified(Ctr) Then
                        NoOfRecords = NoOfRecords + 1
                    End If
                Next
            End If
            ThisForm.Variables("No.Of.RecordModified").SetValue(NoOfRecords)

        End Sub

        Sub SetSupplQtyConvFactor()

            Dim SupplQtyFactor As Double

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CommCode") = "" Then
                SupplQtyFactor = 1.0
            Else
                SupplQtyFactor = ThisForm.Variables("vSupplQtyFactor").GetValueOfDouble(1.0)
            End If

            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefreshInternal(("SupplQtyConvFactor"), Mongoose.Core.Common.MGType.ToInternal(SupplQtyFactor))

        End Sub

        Sub CallDoGeneratePriceCalc()
            If ThisForm.Variables("PrevUM").Value <>
                ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("UM") And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerNonInventoryItemFlag") = "0" Then
                'Non-Inv Item no need to do this.
                ThisForm.Variables("CalculatePrice").Value = "0"
                ThisForm.GenerateEvent("DoGeneratePriceCalculation")

            End If

        End Sub

        Sub ValidateShipSiteEdit()
            If Not ThisForm.PrimaryIDOCollection.IsCurrentObjectNewAndUnmodified Then
                'Required to keep site field from forcing modified flag to true.
                ThisForm.Components("ShipSiteEdit").ValidateData(True)
            End If
        End Sub
        Sub ValidateCoNumEdit()
            'Case when run stand alone, not linked, run validator
            If ThisForm.ParentFormName = "" And
            Trim(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")) <> "" Then
                ThisForm.Components("CoNumEdit").ValidateData(True)
            End If
        End Sub

        Sub ValidateWhseEdit()
            ThisForm.Components("WhseEdit").ValidateData(True)
            ThisForm.GenerateEvent("WhseQtyOrderedValid")
        End Sub

        Sub ChgFeatStr()
            If ThisForm.Variables("FeatStr").Value <> ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr") Then
                ThisForm.GenerateEvent("Reprice")
                ThisForm.CallGlobalScript("FeatStrToGrid", "", "", "",
                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

            End If
        End Sub

        Sub EnableShipSite()
            'ACTION(Enabled:#C(SLCoBlnsShipSiteSubGridCol);#C(ShipSiteEdit), #V(Parm_MultiSite) = 1 & #V(Parm_SharedCust) = 1, True)

            'ShipSiteEdit,ShipSiteGridCol
            If ThisForm.Variables("Parm_MultiSite").Value = "1" And ThisForm.Variables("Parm_SharedCust").Value = "1" Then
                ThisForm.Variables("ShipSiteEnabled").Value = "1"
            Else
                ThisForm.Variables("ShipSiteEnabled").Value = "0"
            End If

        End Sub
        Sub ResetFrtTaxCode()
            Dim strTmp As String
            strTmp = ThisForm.Variables("TaxCodeTypeVar1").Value
            ThisForm.Variables("TaxCodeTypeVar1").Value = ThisForm.Variables("FrtTaxCodeTypeVar1").Value
            ThisForm.Variables("FrtTaxCodeTypeVar1").Value = strTmp
        End Sub


        Sub ReCalcCurrSiteItemPrice()
            Dim bShipSite As Boolean
            bShipSite = (UCase(ThisForm.Variables("PrevSite").Value) <>
                             UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")))
            If bShipSite Then
                'Do not generate pricing event
                ThisForm.Variables("CalculatePrice").Value = "1"
                ThisForm.GenerateEvent("CalculateCoitemPrice")
            End If
        End Sub

        Sub DefaultCoLine()
            ThisForm.Components("CoLineEdit").DefaultData()
        End Sub

        Sub EnableDisableFeaturesTab()

            'ProdConfTab
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItPlanFlag") = "1" Or ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr") <> "" Then
                ThisForm.Variables("ProdConfTabEnabled").SetValue("1")
            Else
                ThisForm.Variables("ProdConfTabEnabled").SetValue("0")
            End If

        End Sub
        Sub ResetCurrentObjectModified()
            ThisForm.PrimaryIDOCollection.SetCurrentObjectModified(True)
        End Sub
        Sub SetQtyOrdAndUM()
            ThisForm.Components("QtyOrderedConvEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("UMEdit").SetModifiedSinceLoadOrValidation(False)
        End Sub

        Sub ResetDerQtyReadyConv()
            If ThisForm.Variables("CopyEvent").GetValue(Of Integer)() = 1 Then ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty(
                    "DerQtyReadyConv", "0")

        End Sub


        Function ChkCustNumIsBlank() As Integer
            ChkCustNumIsBlank = 0
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CustNum") = "" Then ChkCustNumIsBlank = -1
        End Function

        Sub CallCalculateCoItemPrice()
            Dim bCustItemChanged As Boolean
            bCustItemChanged = (ThisForm.Variables("PrevCustItem").Value <>
                             ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CustItem"))
            If bCustItemChanged Then
                ThisForm.Variables("CalculatePrice").Value = "1"
                ThisForm.GenerateEvent("CalculateCoitemPrice")
            End If
        End Sub

        Sub EnableSourceTab()
            Dim RefType As String = ThisForm.Components("RefTypeEdit").Value
            ThisForm.Components("InventoryTab").Enabled = RefType = "I"
            'RefNumStatic
            'RefLineSufStatic
            'RefReleaseStatic
            'RefNumEdit
            'RefLineSufEdit
            'RefReleaseEdit

            ThisForm.Variables("RefNumVisible").SetValue("1")
            ThisForm.Variables("RefLineSufVisible").SetValue("1")
            ThisForm.Variables("RefReleaseVisible").SetValue("1")

            If RefType = "P" Then
                ThisForm.Components("RefNumStatic").Caption = "sPurchaseOrder"
                ThisForm.Components("RefLineSufStatic").Caption = "sLine"
                ThisForm.Components("RefReleaseStatic").Caption = "sRelease"
            ElseIf RefType = "J" Then
                ThisForm.Components("RefNumStatic").Caption = "sJob"
                ThisForm.Components("RefLineSufStatic").Caption = "sSuffix"
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            ElseIf RefType = "R" Then
                ThisForm.Components("RefNumStatic").Caption = "sRequisition"
                ThisForm.Components("RefLineSufStatic").Caption = "sLine"
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            ElseIf RefType = "T" Then
                ThisForm.Components("RefNumStatic").Caption = "sTransfer"
                ThisForm.Components("RefLineSufStatic").Caption = "sLine"
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            ElseIf RefType = "K" Then
                ThisForm.Components("RefNumStatic").Caption = "sProject"
                ThisForm.Components("RefLineSufStatic").Caption = "sLine"
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            ElseIf RefType = "S" And Application.Variables("Avail_FSP").Value = "1" Then
                ThisForm.Components("RefNumStatic").Caption = "sSRO"
                ThisForm.Variables("RefLineSufVisible").SetValue("0")
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            Else
                ThisForm.Variables("RefNumVisible").SetValue("0")
                ThisForm.Variables("RefLineSufVisible").SetValue("0")
                ThisForm.Variables("RefReleaseVisible").SetValue("0")
            End If

        End Sub

        Sub SetConfigPushButton()

            'ConfigPushButton
            If ThisForm.Components("ConfigPushButton").Enabled = True And
               Application.Variables("Avail_Cfg").Value = "1" And
               ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerCfgJobIsConfigurable") = "1" Then
                ThisForm.Variables("ConfigPushEnabled").SetValue("1")
            Else
                ThisForm.Variables("ConfigPushEnabled").SetValue("0")
            End If
        End Sub

        Sub SetSourceSubTabRefresh()

            'If not on the Source tab, don't do this refresh
            If ThisForm.Components("Notebook").NotebookCurTab <> "SourceTab" Then
                Return
            End If

            Select Case ThisForm.Components("RefTypeEdit").Value
                Case "I"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLItemlocAlls", -1).Refresh()
                Case ("P")
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLPoItems", -1).Refresh()
                Case "J"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLJobRoutes", -1).Refresh()
                Case "R"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLPreqitems", -1).Refresh()
                Case "T"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLTrnitems", -1).Refresh()
                Case "K"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("SLProjTasks", -1).Refresh()
                Case "S"
                    ThisForm.PrimaryIDOCollection.GetSubCollection("FSSROs", -1).Refresh()
            End Select


        End Sub


        Sub SetCfgCreateJob()
            If ThisForm.Variables("PromptResponse").Value = CStr(vbYes) Then
                ThisForm.Variables("CfgCreateJob").Value = "1"
            Else
                ThisForm.Variables("CfgCreateJob").Value = "0"
            End If
        End Sub

        Sub SetOldUnitPrice()
            ' Set OldUnitPrice variable from PriceConv
            ThisForm.Variables("OldUnitPrice").Value = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")
        End Sub


        Sub CollectCoNumForStatusChange()

            If ThisForm.Variables("PrevStat").Value() = "P" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "O" _
                AndAlso ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoPortalOrder") = "1" Then
                Dim StrButtonPressed As String

                StrButtonPressed = ThisForm.CallGlobalScript("MsgApp", "Clear", "OK|Cancel", "SuccessFailure",
                            "Q=PortalOrderStatusChangeOKCancel", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

                If StrButtonPressed = "1" Then
                    ThisForm.SetFocus("StatEdit")
                End If
            End If

            If CBool(InStr("OF", (ThisForm.Variables("PrevStat").Value)) > 0) And
               CBool(InStr("C", (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat"))) > 0) _
               Then
                If CBool(InStr(ThisForm.Variables("CoNumList").Value,
                   (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum"))) = 0) _
                   Then
                    ThisForm.Variables("CoNumList").Value =
                       ThisForm.Variables("CoNumList").Value + "," +
                       ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")
                End If
            End If
        End Sub

        Function SetPriceCodeState() As Integer

            Dim StrPrevValue As String
            Dim StrCacheValue As String
            Dim StrButtonPressed As String
            SetPriceCodeState = -1
            StrPrevValue = ThisForm.Variables("PrevPriceCode").Value
            StrCacheValue = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Pricecode")

            If Not StrPrevValue = StrCacheValue Then
                StrButtonPressed = ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                        "mQ=CmdPerform0NoYes", "@sRecalculate", "@sUnitPrice", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                ThisForm.Variables("PrevPriceCode").Value = StrCacheValue
                If StrButtonPressed = "0" Then
                    ThisForm.Variables("CalculatePrice").Value = "0"
                    SetPriceCodeState = 0
                End If
            End If
        End Function


        Sub QueryToChangeCoStatus()

            Dim iDelim1 As Integer
            Dim iDelim2 As Integer

            Dim strOneSet As String

            iDelim1 = 0
            iDelim2 = 0

            If Trim(ThisForm.Variables("CoNumList").Value) <> "" Then
                ThisForm.GenerateEvent("CheckCoStatus")

                If Trim(ThisForm.Variables("CoNumAndStatList").Value) <> "" _
                Then
                    ThisForm.Variables("CoNumAndStatList").Value =
                       (Mid(ThisForm.Variables("CoNumAndStatList").Value,
                        2,
                        Len(ThisForm.Variables("CoNumAndStatList").Value) - 1)) + ","
                End If

                Do While Trim(ThisForm.Variables("CoNumAndStatList").Value) <> ""

                    iDelim1 = InStr(ThisForm.Variables("CoNumAndStatList").Value, ",")
                    strOneSet = Mid(ThisForm.Variables("CoNumAndStatList").Value, 1, (iDelim1 - 1))
                    iDelim2 = InStr(strOneSet, ";")

                    ThisForm.Variables("CoNum").Value = Mid(strOneSet, 1, iDelim2 - 1)
                    ThisForm.Variables("CoStat").Value = Mid(strOneSet, (iDelim2 + 1), (Len(strOneSet) - iDelim2))

                    ThisForm.GenerateEvent("msgCanSetCoStatus")
                    If ThisForm.Variables("PromptResponse").Value = CStr(vbYes) Then
                        ThisForm.Variables("CoNumCanSet").Value = ThisForm.Variables("CoNumCanSet").Value + "," + ThisForm.Variables("CoNum").Value
                    End If

                    ThisForm.Variables("CoNumAndStatList").Value =
                       Mid(
                           ThisForm.Variables("CoNumAndStatList").Value,
                           (iDelim1 + 1),
                           (Len(ThisForm.Variables("CoNumAndStatList").Value) - iDelim1)
                          )
                Loop

                If Trim(ThisForm.Variables("CoNumCanSet").Value) <> "" Then
                    ThisForm.GenerateEvent("SetCoStatus")
                End If

                ThisForm.Variables("CoNumList").Value = ""
                ThisForm.Variables("CoNumAndStatList").Value = ""
                ThisForm.Variables("CoNumCanSet").Value = ""

            End If
        End Sub
        Sub ApplyMask()
            Dim strAmtMask As String
            Dim strCstPrcMask As String

            strAmtMask = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CurrencyAmtFormat")
            strCstPrcMask = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CurrencyCstPrcFormat")

            If Not ThisForm.Components("NetPriceEdit").InputMask.Equals(strAmtMask) Then
                ThisForm.Components("NetPriceEdit").InputMask = strAmtMask
                ThisForm.Components("NetPriceGridCol").InputMask = strAmtMask
                ThisForm.Components("ExtPriceEdit").InputMask = strAmtMask
                ThisForm.Components("ExtPriceGridCol").InputMask = strAmtMask
                ThisForm.Components("ExportValueEdit").InputMask = strAmtMask
                ThisForm.Components("ExportValueGridCol").InputMask = strAmtMask
            End If

            If Not ThisForm.Components("PriceConvEdit").InputMask.Equals(strCstPrcMask) Then
                ThisForm.Components("PriceConvEdit").InputMask = strCstPrcMask
                ThisForm.Components("PriceConvGridCol").InputMask = strCstPrcMask
            End If
        End Sub

        Sub CalculatePrice()
            If ThisForm.PrimaryIDOCollection.IsCurrentObjectNew Then
                ThisForm.GenerateEvent("CalculateCoitemPrice")
            End If
        End Sub

        Sub CheckItemCustXRef()
            Dim iReturn As Integer
            Dim oCache As IWSIDOCollection
            Dim NewItem As String

            oCache = ThisForm.PrimaryIDOCollection
            NewItem = ThisForm.Variables("NewItem").Value
            oCache.SetCurrentObjectProperty("UbItemCustAdd", "0")
            oCache.SetCurrentObjectProperty("UbItemCustUpdate", "0")
            If oCache.GetCurrentObjectProperty("CustItem") = "" Then
                'We're done if CustItem is empty
                Exit Sub
            End If

            If NewItem <> "" And
             UCase(oCache.GetCurrentObjectProperty("Item")) <> UCase(NewItem) And
             EnableIsItemEnabled() Then
                oCache.SetCurrentObjectPropertyPlusModifyRefresh("Item", NewItem)
                If ThisForm.Components("ItemEdit").ValidateData(True) Then
                    'If Validation succeeds, then run the Data changed event handler
                    ThisForm.GenerateEvent("ItemChanged")
                End If
                'May need to do a data changed event also or just fire appropriate events
            ElseIf ThisForm.Variables("ItemCustAdd").Value = "1" Or
               ThisForm.Variables("ItemCustUpdate").Value = "1" Then
                iReturn = ThisForm.GenerateEvent("ItemCustSavePrompt")
                Dim iResponse As wsMsgBoxResult = ThisForm.Variables("PromptResponse").GetValue(Of wsMsgBoxResult)()

                If iReturn = 0 Then
                    'Set appropriate properties for Save time action.
                    If iResponse = vbYes And ThisForm.Variables("ItemCustAdd").Value = "1" Then
                        oCache.SetCurrentObjectProperty("UbItemCustAdd", "1")
                    ElseIf iResponse = vbYes And ThisForm.Variables("ItemCustUpdate").Value = "1" And ThisForm.Variables("ItemCustAdd").Value = "0" Then
                        oCache.SetCurrentObjectProperty("UbItemCustUpdate", "1")
                    ElseIf iResponse = vbNo And ThisForm.Variables("ItemCustUpdate").Value = "1" And ThisForm.Variables("ItemCustAdd").Value = "1" Then
                        oCache.SetCurrentObjectProperty("UbItemCustUpdate", "1")
                    End If
                End If
            End If
            Exit Sub
        End Sub


        Sub SetupConfig()
            Dim dQty As Decimal
            Dim iIndex As Integer

            If IsNumeric(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyOrderedConv")) Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQty = ThisForm.PrimaryIDOCollection(iIndex)("QtyOrderedConv").GetValueOfDecimal(0)
            Else
                dQty = 0
            End If

            If (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItPlanFlag") = "1") Or (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr") <> "") Then
                ThisForm.Variables("ProdConfTabEnabled").SetValue("1")
            Else
                ThisForm.Variables("ProdConfTabEnabled").SetValue("0")
            End If


            If Not (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "P" Or
                    (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "O" And dQty = 0) Or
                    (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr") = "")) Then
                ThisForm.Variables("ItemPlanFlag").Value = "0"
            End If
        End Sub

        Sub EnableInvFreqSummarize()
            Dim loadSLCustomersIDO As LoadCollectionResponseData
            Dim loadMXSATParmsIDO As LoadCollectionResponseData

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Consolidate") = "1" Then
               If Application.Variables("IsCSIB_97762ActiveVar").Value = "1" And Application.Variables("Avail_MX").Value = "1" Then
                  loadSLCustomersIDO = IDOClient.LoadCollection("SLCustomers", "InvFreq,CustNum,CustSeq", "CustNum='" + ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCustNum") + "' AND CustSeq=0", "", 1)
                  loadMXSATParmsIDO = IDOClient.LoadCollection("MXSATParms", "InvFreq,GlobalInvCustNum,GlobalInvCustSeq", "GlobalInvCustNum='" + ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCustNum") + "' AND GlobalInvCustSeq=0", "", 1)
                  If loadSLCustomersIDO.Items.Count = 1 And loadMXSATParmsIDO.Items.Count = 1 Then
                     ThisForm.Variables("InvFreqEnabled").SetValue("0")
                  Else
                     ThisForm.Variables("InvFreqEnabled").SetValue("1")
                  End If
               Else
                  ThisForm.Variables("InvFreqEnabled").SetValue("1")
               End If

               ThisForm.Variables("SummarizeEnabled").SetValue("1")
            Else
                ThisForm.Variables("InvFreqEnabled").SetValue("0")
                ThisForm.Variables("SummarizeEnabled").SetValue("0")
            End If
        End Sub

        Function CallCalcCoItemPriceOnQtyChange() As Integer

            CallCalcCoItemPriceOnQtyChange = 0
            If ThisForm.Variables("CalculatePrice").Value = "0" _
            Or UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item")) <> UCase(ThisForm.Variables("CurrItem").Value) _
            Or ThisForm.PrimaryIDOCollection.IsCurrentObjectNew Then
                'Do not generate pricing event
                ThisForm.Variables("CalculatePrice").Value = "1"
                ThisForm.GenerateEvent("CalculateCoitemPrice")
            End If

        End Function

        Sub EnableFields()
            Dim sStat As String
            Dim bPlanned As Boolean
            Dim bIsNew As Boolean
            Dim bOrdered As Boolean
            Dim dQtyRsvd As Decimal
            Dim dQtyOrd As Decimal
            Dim dQtyShipped As Decimal
            Dim sCOStat As String
            Dim bCanChangeItem As Boolean
            Dim bCanChgShipSite As Boolean
            Dim bComplete As Boolean
            Dim bSetEnabled As Boolean
            Dim bEcReporting As Boolean
            Dim bOrigSite As Boolean
            Dim bShipSite As Boolean
            Dim bTax1Prompt As Boolean
            Dim bTax2Prompt As Boolean
            Dim bSuppl As Boolean
            Dim oCache As IWSIDOCollection
            Dim bEdiOrder As Boolean
            Dim bCanChangeWhse As Boolean
            Dim iIndex As Integer
            Dim bSerialTracked As Boolean
            Dim bJobConfigurable As Boolean
            Dim bRefEnabled As Boolean
            Dim bPromotionCodeNull As Boolean
            Dim bNonInventoryItem As Boolean
            Dim bPostJour As Boolean
            Dim bFeatStrBlank As Boolean
            Dim sFeatStr As String
            Dim bExtShipStatEdit As Boolean
            Dim dIsOnPickList As Decimal
            Dim loadSLCustomersIDO As LoadCollectionResponseData
            Dim loadMXSATParmsIDO As LoadCollectionResponseData

            'Except for Drop Ship tab components, all other components
            'should be enabled only if OrigSite = ParmsSite;
            'Drop Ship Tab components should be enabled only if ShipSite = ParmsSite

            bOrigSite = (UCase(Application.Variables("Parm_Site").Value) =
                         UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite")))
            bShipSite = (UCase(Application.Variables("Parm_Site").Value) =
                         UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")))

            bEcReporting = Application.Variables("Parm_EcReporting").GetValueOfBoolean(False)


            bExtShipStatEdit = (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ExternalShipmentStatus") <> "A")

            bIsNew = ThisForm.PrimaryIDOCollection.IsCurrentObjectNew
            sStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat")
            bPlanned = sStat = "P" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerOldStat") <> "C"
            bOrdered = sStat = "O" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerOldStat") <> "C"
            bComplete = sStat = "C" Or
                      sStat = "H"
            bPromotionCodeNull = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PromotionCode") = ""
            bNonInventoryItem = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerNonInventoryItemFlag") = "1"
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyOrderedConv") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQtyOrd = ThisForm.PrimaryIDOCollection(iIndex)("QtyOrderedConv").GetValueOfDecimal(0)
            Else
                dQtyOrd = 0
            End If

                 dQtyRsvd = ThisForm.PrimaryIDOCollection.CurrentItem("QtyRsvd").GetValueOfDecimal(0)

              dQtyShipped = ThisForm.PrimaryIDOCollection.CurrentItem("DerQtyShippedConv").GetValueOfDecimal(0)

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ComSupplQtyReq") = "1" Then
                bSuppl = True
            Else
                bSuppl = False
            End If
            sCOStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoStat")
            bCanChgShipSite = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerCanChgShipSite") = "1"
            bSerialTracked = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItSerialTracked") = "1"
            bJobConfigurable = Not ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerCfgJobIsConfigurable") = "0"
            sFeatStr = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr")
            If Trim(Replace(sFeatStr, ",", "")) = "" Or Trim(Replace(sFeatStr, ",", "")) = "." Then
                bFeatStrBlank = True 'sFbomBlank
            Else
                bFeatStrBlank = False
            End If
            ' set functions

            Dim SaveAtEnabledGetATPCTPFlag As String
            SaveAtEnabledGetATPCTPFlag = ThisForm.Variables("SaveAtEnabledGetATPCTPFlagVar").Value

             dIsOnPickList = ThisForm.PrimaryIDOCollection.CurrentItem("DerIsOnPickList").GetValueOfDecimal(0)

            bCanChangeItem = bIsNew Or (Not bIsNew And bPlanned Or bOrdered And dQtyOrd = 0 And dQtyRsvd = 0)
            bCanChangeWhse = (dQtyRsvd = 0) And bCanChangeItem And (dIsOnPickList = 0)
            '  plus has no releases for bln orders

            'Enabling components
            'CustNumEdit,CustNumGridCol,CustSeqEdit,CustSeqGridCol
            If (sStat <> "F" And sStat <> "C") And bShipSite Then
                ThisForm.Variables("CustNumEnabled").SetValue("1")
                ThisForm.Variables("CustSeqEnabled").SetValue("1")
            Else
                ThisForm.Variables("CustNumEnabled").SetValue("0")
                ThisForm.Variables("CustSeqEnabled").SetValue("0")
            End If

            'ItemDescEdit,ItemDescGridCol,CustItemEdit,CustItemGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite Then
                ThisForm.Variables("ItemDescEnabled").SetValue("1")
                ThisForm.Variables("CustItemEnabled").SetValue("1")
            Else
                ThisForm.Variables("ItemDescEnabled").SetValue("0")
                ThisForm.Variables("CustItemEnabled").SetValue("0")
            End If

            'ItemEdit,ItemGridCol
            If bCanChangeItem And bOrigSite And bExtShipStatEdit Then
                ThisForm.Variables("ItemEnabled").SetValue("1")
            Else
                ThisForm.Variables("ItemEnabled").SetValue("0")
            End If

            'QtyOrderedConvEdit,QtyOrderedConvGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And bExtShipStatEdit Then
                ThisForm.Variables("QtyOrderedConvEnabled").SetValue("1")
            Else
                ThisForm.Variables("QtyOrderedConvEnabled").SetValue("0")
            End If

            'UMEdit,UMGridCol
            If (bPlanned Or bOrdered Or bIsNew) And dQtyShipped = 0 And bOrigSite And Not bSerialTracked And bExtShipStatEdit Then
                ThisForm.Variables("UMEnabled").SetValue("1")
            Else
                ThisForm.Variables("UMEnabled").SetValue("0")
            End If

            'PriceConvEdit,PriceConvGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And bPromotionCodeNull Then
                ThisForm.Variables("PriceConvEnabled").SetValue("1")
            Else
                ThisForm.Variables("PriceConvEnabled").SetValue("0")
            End If

            'RepriceBtn
            If Application.Variables("Avail_AU").Value = "1" Then
                If (bPlanned Or bOrdered Or sStat = "F" Or sStat = "C" Or bIsNew) And bOrigSite And Not bNonInventoryItem Then
                    ThisForm.Variables("RepriceEnabled").SetValue("1")
                Else
                    ThisForm.Variables("RepriceEnabled").SetValue("0")
                End If
            Else
                If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And Not bNonInventoryItem Then
                    ThisForm.Variables("RepriceEnabled").SetValue("1")
                Else
                    ThisForm.Variables("RepriceEnabled").SetValue("0")
                End If
            End If


            'DueDateEdit,DueDateGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And bExtShipStatEdit Then
                ThisForm.Variables("DueDateEnabled").SetValue("1")
            Else
                ThisForm.Variables("DueDateEnabled").SetValue("0")
            End If

            'PromiseDateEdit,PromiseDateGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite Then
                ThisForm.Variables("PromiseDateEnabled").SetValue("1")
            Else
                ThisForm.Variables("PromiseDateEnabled").SetValue("0")
            End If

            'AUDueDateEdit,AUDueDateGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And bExtShipStatEdit Then
                ThisForm.Variables("AUDueDateEnabled").SetValue("1")
            Else
                ThisForm.Variables("AUDueDateEnabled").SetValue("0")
            End If

            'AUPromiseDateEdit,AUPromiseDateGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite Then
                ThisForm.Variables("AUPromiseDateEnabled").SetValue("1")
            Else
                ThisForm.Variables("AUPromiseDateEnabled").SetValue("0")
            End If

            ' Is this a Buy Design Configuration and is the Job Released, then disable the Ref Fields.
            bRefEnabled = (bPlanned Or bOrdered Or bIsNew) And (bShipSite Or bIsNew)

            'RefTypeEdit,RefTypeGridCol,RefNumEdit,RefNumGridCol,RefLineSufEdit,RefLineSufGridCol,RefReleaseEdit,RefReleaseGridCol
            If bRefEnabled Then
                ThisForm.Variables("RefTypeEnabled").SetValue("1")
                ThisForm.Variables("RefNumEnabled").SetValue("1")
                ThisForm.Variables("RefLineSufEnabled").SetValue("1")
                ThisForm.Variables("RefReleaseEnabled").SetValue("1")
            Else
                ThisForm.Variables("RefTypeEnabled").SetValue("0")
                ThisForm.Variables("RefNumEnabled").SetValue("0")
                ThisForm.Variables("RefLineSufEnabled").SetValue("0")
                ThisForm.Variables("RefReleaseEnabled").SetValue("0")
            End If

            'WhseEdit,WhseGridCol
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoConsignment") = "0" And bCanChangeWhse And bOrigSite Then
                ThisForm.Variables("WhseEnabled").SetValue("1")
            Else
                ThisForm.Variables("WhseEnabled").SetValue("0")
            End If

            bTax1Prompt = Application.Variables("TaxP_PromptForSystem1").Value = "1" And
                          Application.Variables("Tax1_PromptOnLine").Value = "1"

            'TaxCode1Edit,TaxCode1GridCol
            If (Not bComplete Or bIsNew) And bOrigSite And bTax1Prompt Then
                ThisForm.Variables("TaxCode1Enabled").SetValue("1")
            Else
                ThisForm.Variables("TaxCode1Enabled").SetValue("0")
            End If

            If Application.Variables("Tax2_Enabled").Value = "1" Then
                bTax2Prompt = Application.Variables("TaxP_PromptForSystem2").Value = "1" And
                              Application.Variables("Tax2_PromptOnLine").Value = "1"

                'TaxCode2Edit,TaxCode2GridCol
                If (Not bComplete Or bIsNew) And bOrigSite And bTax2Prompt Then
                    ThisForm.Variables("TaxCode2Enabled").SetValue("1")
                Else
                    ThisForm.Variables("TaxCode2Enabled").SetValue("0")
                End If
            End If

            bEdiOrder = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoEdiOrder") = "1"
            'ConsolidateEdit,ConsolidateGridCol
            If Not bEdiOrder And (bPlanned Or bOrdered Or bIsNew) Then
                ThisForm.Variables("ConsolidateEnabled").SetValue("1")
            Else
                ThisForm.Variables("ConsolidateEnabled").SetValue("0")
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Consolidate") = "1" Then
                bSetEnabled = True
                If bEdiOrder Then
                    ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Consolidate", "0")
                    ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Summarize", "0")
                End If
            Else
                bSetEnabled = False
            End If

            'InvFreqEdit,InvFreqGridCol,SummarizeEdit,SummarizeGridCol
            If bSetEnabled Then
               If Application.Variables("IsCSIB_97762ActiveVar").Value = "1" And Application.Variables("Avail_MX").Value = "1" Then
                  loadSLCustomersIDO = IDOClient.LoadCollection("SLCustomers", "InvFreq,CustNum,CustSeq", "CustNum='" + ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCustNum") + "' AND CustSeq=0", "", 1)
                  loadMXSATParmsIDO = IDOClient.LoadCollection("MXSATParms", "InvFreq,GlobalInvCustNum,GlobalInvCustSeq", "GlobalInvCustNum='" + ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCustNum") + "' AND GlobalInvCustSeq=0", "", 1)
                  If loadSLCustomersIDO.Items.Count = 1 And loadMXSATParmsIDO.Items.Count = 1 Then
                     ThisForm.Variables("InvFreqEnabled").SetValue("0")
                  Else
                     ThisForm.Variables("InvFreqEnabled").SetValue("1")
                  End If
               Else
                  ThisForm.Variables("InvFreqEnabled").SetValue("1")
               End If

               ThisForm.Variables("SummarizeEnabled").SetValue("1")
            Else
                ThisForm.Variables("InvFreqEnabled").SetValue("0")
                ThisForm.Variables("SummarizeEnabled").SetValue("0")
            End If

            'FeatStrEdit,FeatStrGridCol,FMSelectedGridCol,FMMatlQtyGridCol,FMSelectedViewGridCol,FMMatlQtyViewGridCol
            If (((bPlanned Or bFeatStrBlank) And Not bComplete) Or bIsNew) And bOrigSite And bJobConfigurable Then
                ThisForm.Variables("FeatStrEnabled").SetValue("1")
                ThisForm.Variables("FMSelectedEnabled").SetValue("1")
                ThisForm.Variables("FMMatlQtyEnabled").SetValue("1")
                ThisForm.Variables("FMSelectedViewEnabled").SetValue("1")
                ThisForm.Variables("FMMatlQtyViewEnabled").SetValue("1")
            Else
                ThisForm.Variables("FeatStrEnabled").SetValue("0")
                ThisForm.Variables("FMSelectedEnabled").SetValue("0")
                ThisForm.Variables("FMMatlQtyEnabled").SetValue("0")
                ThisForm.Variables("FMSelectedViewEnabled").SetValue("0")
                ThisForm.Variables("FMMatlQtyViewEnabled").SetValue("0")
            End If

            'PricecodeEdit,PricecodeGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite Then
                ThisForm.Variables("PricecodeEnabled").SetValue("1")
            Else
                ThisForm.Variables("PricecodeEnabled").SetValue("0")
            End If

            'DiscEdit,DiscGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And bPromotionCodeNull Then
                ThisForm.Variables("DiscEnabled").SetValue("1")
            Else
                ThisForm.Variables("DiscEnabled").SetValue("0")
            End If

            'ECVAT Tab
            ' This form variable is used in Component Conditional Actions for EU VAT components
            If (Not bComplete Or bIsNew) And bEcReporting And bOrigSite Then
                ThisForm.Variables("EnableEUVATVar").Value = "1"
            Else
                ThisForm.Variables("EnableEUVATVar").Value = "0"
            End If

            Call EnableDisableSupplQtyConvFactor()

            oCache = ThisForm.PrimaryIDOCollection
            If oCache.GetCurrentObjectProperty("DerCanChgShipSite") = "1" And
                ThisForm.Variables("SharedCustEnabled").Value = "1" And
                EnableIsItemEnabled() And
                 (UCase(Application.Variables("Parm_Site").Value) =
                         UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite"))) Then

                ThisForm.Variables("MultiSiteItemSourcingEnabled").SetValue("1")
                ThisForm.Variables("ShipSiteEnabled").SetValue("1")
            Else
                ThisForm.Variables("MultiSiteItemSourcingEnabled").SetValue("0")
                ThisForm.Variables("ShipSiteEnabled").SetValue("0")
            End If

            If (sCOStat = "C" Or (Not bOrigSite)) And ThisForm.ParentFormName <> "" Then
                ThisForm.PrimaryIDOCollection.NewEnabled = False
            Else
                ThisForm.PrimaryIDOCollection.NewEnabled = True
            End If
            If sCOStat = "C" Then
                ThisForm.PrimaryIDOCollection.DeleteEnabled = False
                ThisForm.PrimaryIDOCollection.SaveEnabled = False
            Else
                ThisForm.PrimaryIDOCollection.DeleteEnabled = bOrigSite
                ThisForm.PrimaryIDOCollection.SaveEnabled = True
            End If

            ThisForm.Variables("PrevPriceCode").Value = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Pricecode")
            bPostJour = ThisForm.Variables("PostJourVar").Value = "1"
            ' This form variable is used in Component Conditional Actions for Non Inventory Acct components
            If (bNonInventoryItem Or (Not bNonInventoryItem And Not bPostJour)) And (sStat <> "F" And sStat <> "C") And
                ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item") <> "" Then
                ThisForm.Variables("EnableNonInvAcctCompsVar").Value = "1"
            Else
                ThisForm.Variables("EnableNonInvAcctCompsVar").Value = "0"
            End If

            'CostConvEdit,CostConvGridCol
            If bNonInventoryItem And dQtyShipped = 0 And (sStat <> "F" And sStat <> "C") Then
                ThisForm.Variables("CostConvEnabled").SetValue("1")
            Else
                ThisForm.Variables("CostConvEnabled").SetValue("0")
            End If

            'DerTotCostEdit,DerTotCostGridCol
            If bNonInventoryItem And (sStat <> "F" And sStat <> "C") Then
                ThisForm.Variables("DerTotCostEnabled").SetValue("1")
            Else
                ThisForm.Variables("DerTotCostEnabled").SetValue("0")
            End If


            'PromotionCodeEdit,PromotionCodeGridCol
            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite And Not bNonInventoryItem _
                And Not (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItPlanFlag") = "1") _
                And Not (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItOrderConfigurable") = "1") Then
                ThisForm.Variables("PromotionCodeEnabled").SetValue("1")
                Call RefreshPromotionCodes()
            Else
                ThisForm.Variables("PromotionCodeEnabled").SetValue("0")
            End If

            'Disable the fields related to PO-CO. Remove the calling in SelectCurrentCompleted() and move it here.
            EnableDemandingSite()
        End Sub

        Sub EnableDisableSupplQtyConvFactor()

            Dim sStat As String
            Dim bComplete As Boolean
            Dim bIsNew As Boolean
            Dim bEcReporting As Boolean
            Dim bOrigSite As Boolean
            Dim bSuppl As Boolean

            sStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat")
            bComplete = sStat = "C" Or sStat = "H"
            bIsNew = ThisForm.PrimaryIDOCollection.IsCurrentObjectNew
            bEcReporting = Application.Variables("Parm_EcReporting").GetValueOfBoolean(False)
            bOrigSite = (UCase(Application.Variables("Parm_Site").Value) = UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite")))
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ComSupplQtyReq") = "1" Then
                bSuppl = True
            Else
                bSuppl = False
            End If

            'SupplQtyConvFactorEdit,SupplQtyConvFactorGridCol
            If (Not bComplete Or bIsNew) And bEcReporting And bOrigSite And bSuppl Then
                ThisForm.Variables("SupplQtyConvFactorEnabled").SetValue("1")
            Else
                ThisForm.Variables("SupplQtyConvFactorEnabled").SetValue("0")
            End If

        End Sub




        Sub EnableXrefButton()
            'XRefButton
            If (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "C" And
            ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefNum") = "") Then
                ThisForm.Variables("XRefButtonEnabled").SetValue("0")
                Exit Sub
            End If

            If UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")) <> UCase(Application.Variables("Parm_Site").Value) Or
               ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCreditHold") = "1" Or ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("AdrCreditHold") = "1" Then
                ThisForm.GenerateEvent("IsCoitemXreferenced")
                If ThisForm.Variables("Xreferenced").Value = "1" Then
                    ThisForm.Variables("XRefButtonEnabled").SetValue("1")
                Else
                    ThisForm.Variables("XRefButtonEnabled").SetValue("0")
                End If
            Else
                If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCreditHold") = "1" _
                Or ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("AdrCreditHold") = "1" Then
                    ThisForm.GenerateEvent("IsCoitemXreferenced")
                    If ThisForm.Variables("Xreferenced").Value = "1" Then
                        ThisForm.Variables("XRefButtonEnabled").SetValue("1")
                    Else
                        ThisForm.Variables("XRefButtonEnabled").SetValue("0")
                    End If
                End If
                ThisForm.Variables("XRefButtonEnabled").SetValue("1")
            End If
        End Sub
        Sub EnterFilterInPlace()

            ThisForm.Variables("RefNumVisible").SetValue("1")
            ThisForm.Variables("RefLineSufVisible").SetValue("1")
            ThisForm.Variables("RefReleaseVisible").SetValue("1")

            ThisForm.Components("RefNumStatic").Caption = "sNumber"
            ThisForm.Components("RefLineSufStatic").Caption = "sLine/Suffix"
            ThisForm.Components("RefReleaseStatic").Caption = "sRelease"
        End Sub

        Sub DefaultCompCaptions()

            If Application.Variables("Tax1_Enabled").Value = "1" Then
                ThisForm.Components("TaxCode1Static").Caption =
                    Application.Variables("Tax1_TaxItemLabel").Value
                ThisForm.Components("TaxCode1GridCol").Caption =
                    Application.Variables("Tax1_TaxItemLabel").Value
                ThisForm.Components("TaxCode1DescGridCol").Caption =
                    Application.Variables("Tax1_TaxItemDescLabel").Value
            End If
            If Application.Variables("Tax2_Enabled").Value = "1" Then
                ThisForm.Components("TaxCode2Static").Caption =
                    Application.Variables("Tax2_TaxItemLabel").Value
                ThisForm.Components("TaxCode2GridCol").Caption =
                    Application.Variables("Tax2_TaxItemLabel").Value
                ThisForm.Components("TaxCode2DescGridCol").Caption =
                    Application.Variables("Tax2_TaxItemDescLabel").Value
            End If

            If ThisForm.Variables("ApsParmApsmode").Value = "M" Then
                ThisForm.Components("GetCTPButton").Caption = Application.GetStringValue("s&GetATP")
            Else
                ThisForm.Components("GetCTPButton").Caption = Application.GetStringValue("s&GetCTP")
            End If
            'Ensure CoTypeVar is set
            If ThisForm.Variables("CoTypeVar").Value = "" Then
                ThisForm.Variables("CoTypeVar").Value = "R"
            End If

        End Sub

        Function EnableIsItemEnabled() As Boolean
            Dim oCache As IWSIDOCollection

            oCache = ThisForm.PrimaryIDOCollection
            If oCache.IsCurrentObjectNew Or
             (UCase(oCache.GetCurrentObjectProperty("CoOrigSite")) =
                 UCase(Application.Variables("Parm_Site").Value) And
              oCache.GetCurrentObjectProperty("DerCanChangeItem") = "1") Then
                EnableIsItemEnabled = True
            Else
                EnableIsItemEnabled = False
            End If
        End Function

        Sub ChkLastModalChildName()
            Dim oCache As IWSIDOCollection
            If ThisForm.LastModalChildEndedOk = True Then
                'if the user hit cancel on the modal form,
                'there is no need to continue.
                If ThisForm.LastModalChildName = "PromptForWarehouseData" Then
                    ThisForm.Variables("ToWhse").Value = ThisForm.ModalChildForm.Variables("ToWhse").Value
                    ThisForm.Variables("FromWhse").Value = ThisForm.ModalChildForm.Variables("FromWhse").Value
                    ThisForm.Variables("ToSite").Value = ThisForm.ModalChildForm.Variables("ToSite").Value
                    ThisForm.Variables("FromSite").Value = ThisForm.ModalChildForm.Variables("FromSite").Value
                    ThisForm.Variables("TrnLoc").Value = ThisForm.ModalChildForm.Variables("TrnLoc").Value
                    ThisForm.Variables("FOBSite").Value = ThisForm.ModalChildForm.Variables("FOBSite").Value

                    ThisForm.Variables("CurRefNum").Value = ThisForm.ModalChildForm.Variables("TrnNum").Value

                    ThisForm.GenerateEvent("CreateTransferOrder")
                End If
                If ThisForm.LastModalChildName = "ShipTosQuery" Then
                    oCache = ThisForm.ModalChildForm.PrimaryIDOCollection
                    If Not ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CustNum").Equals(oCache.GetCurrentObjectProperty("CustNum")) Then
                        ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("CustNum", oCache.GetCurrentObjectProperty("CustNum"))
                        ThisForm.Components("CustNumEdit").ValidateData(True)
                        ThisForm.GenerateEvent("DropShipChanged")
                    End If
                    If Not ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CustSeq").Equals(oCache.GetCurrentObjectProperty("CustSeq")) Then
                        ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("CustSeq", oCache.GetCurrentObjectProperty("CustSeq"))
                        ThisForm.Components("CustSeqEdit").ValidateData(True)
                        ThisForm.GenerateEvent("GetDropShip")
                    End If
                End If
            End If

            ' clean up NextKeys table for failed X-Ref
            If ThisForm.LastModalChildName = "PromptForWarehouseData" And
               ThisForm.Variables("LastKeyRefNum").Value <> "" Then

                ThisForm.GenerateEvent("CleanUpTransNextKey")
                ' clean up RefNum display, when modal has cancled
                If ThisForm.LastModalChildEndedOk = False Or ThisForm.Variables("IsCancel").Value = "1" Then
                    ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(
                                     "RefNum", "")

                End If
            End If

            If ThisForm.LastModalChildName = "ConfigControl" Then
                ThisForm.PrimaryIDOCollection.RefreshKeepCurIndex()
            End If

        End Sub

        Sub CalculateNetPrice()
            Dim iDecimalPlaces As Integer
            Dim dDisc As Decimal
            Dim dPriceConv As Decimal
            Dim dQtyOrd As Decimal
            Dim dNetPrice As Decimal
            Dim iIndex As Integer
            Dim dTweak As Decimal

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Disc") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dDisc = ThisForm.PrimaryIDOCollection(iIndex)("Disc").GetValueOfDecimal(0)
            Else
                dDisc = 0
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dPriceConv = ThisForm.PrimaryIDOCollection(iIndex)("PriceConv").GetValueOfDecimal(0)
            Else
                dPriceConv = 0
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyOrderedConv") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQtyOrd = ThisForm.PrimaryIDOCollection(iIndex)("QtyOrderedConv").GetValueOfDecimal(0)
            Else
                dQtyOrd = 0
            End If

            If ThisForm.Variables("CurrCodePlaces").Value <> "" Then
                iDecimalPlaces = ThisForm.Variables("CurrCodePlaces").GetValueOfInt32(0)
            Else
                iDecimalPlaces = 2
            End If
            dTweak = CDec(1 / (10 ^ (iDecimalPlaces + 8)))

            If ThisForm.Variables("UseAltPriceCalc").Value = "1" Then
                dNetPrice = Round((((1 - dDisc / 100) * dPriceConv) + dTweak), iDecimalPlaces) * dQtyOrd
            Else
                dNetPrice = Round(((dQtyOrd * (1 - dDisc / 100) * dPriceConv) + dTweak), iDecimalPlaces)
            End If

            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DerNetPrice", CStr(dNetPrice))

        End Sub

        Function PreDeleteValidation() As Integer
            Dim sCOStat As String
            Dim sParam As String
            Dim sCoNum As String
            sParam = ""
            PreDeleteValidation = 0

            If ThisForm.PrimaryIDOCollection.IsCurrentObjectDeleted Then
                sCOStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoStat")
                sCoNum = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")

                If sCOStat = "C" Or sCOStat = "H" Then
                    Select Case sCOStat
                        Case "C"
                            sParam = Application.GetStringValue("sCoStatus=C")
                        Case "H"
                            sParam = Application.GetStringValue("sCoStatus=H")
                    End Select

                    ThisForm.CallGlobalScript("MsgApp", "Clear", "NoPrompt", "SuccessFailure",
                        "mE=CmdInvalid", "@sDelete", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                    ThisForm.CallGlobalScript("MsgApp", "NoClear", "Prompt", "SuccessFailure",
                        "mI=IsCompare1", "@sStatus", sParam, "@sCO", "@sOrder", sCoNum, "", "", "", "", "", "", "", "", "", "", "")
                    PreDeleteValidation = -1
                End If
            End If
        End Function

        Function PreSaveValidation() As Integer
            Dim i As Integer

            PreSaveValidation = 0

            For i = 0 To ThisForm.PrimaryIDOCollection.GetNumEntries - 1
                If ThisForm.PrimaryIDOCollection.IsObjectModified(i) _
                    And ThisForm.PrimaryIDOCollection.GetObjectProperty("CoStat", i) = "C" Then
                    ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                        "mI=IsCompare1", "@sStatus", "@sCoStatus=C", "@sCO", "@sOrder", ThisForm.PrimaryIDOCollection.GetObjectProperty("CoNum", i),
                        "", "", "", "", "", "", "", "", "", "", "")

                    PreSaveValidation = -1
                    Exit Function
                ElseIF ThisForm.PrimaryIDOCollection.IsObjectNew(i) and (UCase(Application.Variables("Parm_Site").Value) <> UCase(ThisForm.PrimaryIDOCollection.GetObjectProperty("CoOrigSite", i))) THEN
                       ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                        "mI=IsCompare1", "@sOrigSite", ThisForm.PrimaryIDOCollection.GetObjectProperty("CoOrigSite", i), "@sCO", "@sOrder", ThisForm.PrimaryIDOCollection.GetObjectProperty("CoNum", i),
                        "", "", "", "", "", "", "", "", "", "", "")
                        PreSaveValidation = -1
                    Exit Function
                End If
            Next i
        End Function

        Sub ShowMessagePriceCode()
            ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
              "mQ=CmdPerform-NoYes", "@ssRecalculate", "sUnitPrice", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End Sub

        Sub RefreshRefType()
            Dim sRefType As String
            Dim sOldRefType As String

            sRefType = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefType")
            sOldRefType = ThisForm.Variables("OldRefType").Value

            If sOldRefType <> sRefType Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("RefNum", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("RefLineSuf", "0")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("RefRelease", "0")
            End If
        End Sub

        Sub SetNewRecordFlag()
            If ThisForm.PrimaryIDOCollection.IsCurrentObjectNew And ThisForm.Variables("CalculatePrice").Value = "1" Then
                ThisForm.Variables("NewRecordFlag").Value = "1"
            Else
                ThisForm.Variables("NewRecordFlag").Value = "0"
            End If
        End Sub

        Public Sub SetOrderStatus()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoStat") = "P" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Stat", "P")
            Else
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Stat", "O")
            End If
        End Sub

        Sub SetShipSiteEnabled()
            Dim dQtyOrd As Decimal
            Dim iIndex As Integer
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyOrderedConv") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQtyOrd = ThisForm.PrimaryIDOCollection(iIndex)("QtyOrderedConv").GetValueOfDecimal(0)
            Else
                dQtyOrd = 0
            End If
            If dQtyOrd = 0 Then
                ThisForm.Variables("ShipSiteEnabled").SetValue("1")
            Else
                ThisForm.Variables("ShipSiteEnabled").SetValue("0")
            End If
        End Sub

        Function ValidateCustItemChanged() As Integer
            Dim oCache As IWSIDOCollection

            ValidateCustItemChanged = 0
            oCache = ThisForm.PrimaryIDOCollection
            If UCase(oCache.GetCurrentObjectProperty("CustItem")) <>
              UCase(ThisForm.Variables("PrevCustItem").Value) Then
                If oCache.GetCurrentObjectProperty("CustItem") = "" Then
                    ThisForm.GenerateEvent("CustItemCleared")
                    oCache.SetCurrentObjectProperty("UbItemCustAdd", "0")
                    oCache.SetCurrentObjectProperty("UbItemCustUpdate", "0")

                    ValidateCustItemChanged = 1
                End If
            Else
                ValidateCustItemChanged = 1
            End If
        End Function

        Sub SetTaxCode()
            Dim sTaxCode1 As String
            Dim sTaxCode2 As String

            If ThisForm.Variables("TaxPromptResponce").Value = CStr(vbYes) Then
                sTaxCode1 = ThisForm.Variables("TaxCode1").Value
                sTaxCode2 = ThisForm.Variables("TaxCode2").Value
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("TaxCode1", sTaxCode1)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("TaxCode2", sTaxCode2)
                'ThisForm.Components("TaxCode1Edit").ValidateData True
                'ThisForm.Components("TaxCode1GridCol").ValidateData True
                'ThisForm.Components("TaxCode2Edit").ValidateData True
                'ThisForm.Components("TaxCode2GridCol").ValidateData True
                'Else
                'ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyModified "TaxCode1", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoTaxCode1")
                'ThisForm.Components("TaxCode1Edit").ValidateData True
            End If
        End Sub

        Function ValidateUMChanged() As Integer
            Dim oCache As IWSIDOCollection

            ValidateUMChanged = 0
            oCache = ThisForm.PrimaryIDOCollection
            If UCase(oCache.GetCurrentObjectProperty("UM")) =
             UCase(ThisForm.Variables("PrevUM").Value) Then
                'Nothing more to do
                ValidateUMChanged = 1
            End If
        End Function

        Function Xref() As Integer
            Dim bCreateFlag As Boolean
            Dim bReturnUserXrefAction As Boolean
            Dim strXrefDestination As String
            Dim strFrmRefType As String
            Dim strFrmRefNum As String
            Dim strFrmRefLineSuf As String
            Dim strFrmRefRelease As String
            Dim ReturnCode As Integer

            bCreateFlag = ThisForm.Variables("CreateFlag").Value = "1" _
               Or ThisForm.Variables("CreateFlag2").Value = "1"
            bReturnUserXrefAction = ThisForm.Variables("ReturnUserXrefAction").Value = CStr(vbYes)
            strXrefDestination = ThisForm.Variables("XrefDestination").Value
            strFrmRefType = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefType")
            strFrmRefNum = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefNum")
            strFrmRefLineSuf = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefLineSuf")
            strFrmRefRelease = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefRelease")

            If Not bCreateFlag And strXrefDestination <> "" Then
                If strXrefDestination = "Items" Then
                    ThisForm.Variables("XrefForm").Value = "Items( FILTER(Item='P(Item)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh) )"
                ElseIf strXrefDestination = "JobOrders" Then
                    If ThisForm.Components("RefLineSufEdit").Value = "" Then
                        ThisForm.Components("RefLineSufEdit").Value = "0"
                    End If
                    ThisForm.Variables("XrefForm").Value = "JobOrders( FILTER(Job='P(RefNum)' and Suffix = 'P(RefLineSuf)') OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh) )"
                ElseIf strXrefDestination = "PurchaseOrderBlanketReleases" Then
                    ThisForm.Variables("XrefForm").Value = "PurchaseOrderBlanketReleases(FILTER(PoNum='P(RefNum)'and  PoLine='P(RefLineSuf)' and PoRelease='P(RefRelease)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh))"
                ElseIf strXrefDestination = "PurchaseOrderLines" Then
                    ThisForm.Variables("XrefForm").Value = "PurchaseOrderLines( FILTER(PoNum ='P(RefNum)'  and  PoLine='P(RefLineSuf)') OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh) )"
                ElseIf strXrefDestination = "PurchaseOrderRequisitionLines" Then
                    ThisForm.Variables("XrefForm").Value = "PurchaseOrderRequisitionLines( FILTER(ReqNum='P(RefNum)' and ReqLine='P(RefLineSuf)') OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh) )"
                ElseIf strXrefDestination = "TransferOrderLineItems" Then
                    ThisForm.Variables("XrefForm").Value = "TransferOrderLineItems(FILTER(TrnNum='P(RefNum)' and TrnLine='P(RefLineSuf)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh))"
                ElseIf strXrefDestination = "ProjectTasks" Then
                    ThisForm.Variables("XrefForm").Value = "ProjectTasks(FILTER(ProjNum='P(RefNum)' and TaskNum='P(RefLineSuf)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh))"
                ElseIf strXrefDestination = "SROs" Then
                    ThisForm.Variables("XrefForm").Value = "ServiceOrders(FILTER(SroNum='P(RefNum)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh))"
                ElseIf strXrefDestination = "Incidents" Then
                    ThisForm.Variables("XrefForm").Value = "Incidents(FILTER(IncNum='P(RefNum)')OKCANCELOPTIONAL() SETVARVALUES(InitialCommand=Refresh))"
                End If

                ReturnCode = ThisForm.GenerateEvent("OpenXrefForm")

            ElseIf (bReturnUserXrefAction And
                      (bCreateFlag Or (Not bCreateFlag And strXrefDestination = ""))) Then

                ReturnCode = ThisForm.GenerateEvent("XrefCreate")
            End If

            Xref = ReturnCode
        End Function

        Sub EnableXrefBtnForRefTypeChg()
            If UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")) = UCase(Application.Variables("Parm_Site").Value) And
                (ThisForm.PrimaryIDOCollection.IsCurrentObjectPropertyModified("RefType") Or
                ThisForm.PrimaryIDOCollection.IsCurrentObjectPropertyModified("RefNum") Or
                ThisForm.PrimaryIDOCollection.IsCurrentObjectPropertyModified("RefLineSuf") Or
                ThisForm.PrimaryIDOCollection.IsCurrentObjectPropertyModified("RefRelease") Or
                ThisForm.PrimaryIDOCollection.IsCurrentObjectNew Or
                ThisForm.PrimaryIDOCollection.IsCurrentObjectDeleted) Then
                ThisForm.Variables("XRefButtonEnabled").SetValue("1")
            Else
                ThisForm.Variables("XRefButtonEnabled").SetValue("0")
            End If
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoCreditHold") = "1" Or
               ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("AdrCreditHold") = "1" Then
                ThisForm.Variables("XRefButtonEnabled").SetValue("0")
            End If
        End Sub

        Function XrefCreate() As Integer
            Dim ReturnCode As Integer
            Dim bMpwXrefUserAction As Boolean
            Dim bPoAbortUserAction As Boolean
            Dim bPoChangeOrdUserAction As Boolean
            Dim strPromptMsg1 As String
            Dim strPromptMsg2 As String
            Dim strPromptMsg3 As String

            ReturnCode = 0

            bMpwXrefUserAction = ThisForm.Variables("MpwXrefUserAction").Value = CStr(vbYes)
            bPoAbortUserAction = ThisForm.Variables("PoAbortUserAction").Value = CStr(vbYes)
            bPoChangeOrdUserAction = ThisForm.Variables("PoChangeOrdUserAction").Value = CStr(vbYes)
            strPromptMsg1 = ThisForm.Variables("PromptMsg1").Value
            strPromptMsg2 = ThisForm.Variables("PromptMsg2").Value
            strPromptMsg3 = ThisForm.Variables("PromptMsg3").Value

            If strPromptMsg1 <> "" And bMpwXrefUserAction Then
                ThisForm.Variables("MpwxrefDelete").Value = "1"
            Else
                ThisForm.Variables("MpwxrefDelete").Value = "0"
            End If

            If strPromptMsg2 <> "" And bPoAbortUserAction Then
                ReturnCode = 1
            End If

            If strPromptMsg3 <> "" And bPoChangeOrdUserAction Then
                ThisForm.Variables("PoChangeOrd").Value = "1"
            End If
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefType") = "T" _
             And ThisForm.Variables("CreateFlag").Value = "1" Then
                ThisForm.GenerateEvent("PromptForWhse")
                ReturnCode = 1
            End If
            XrefCreate = ReturnCode
        End Function


        Sub SetPriceConv()
            If Not IsNumeric(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")) Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("PriceConv", "0")
            End If
            If Not IsNumeric(ThisForm.Variables("OldUnitPrice").Value) Then
                ThisForm.Variables("OldUnitPrice").Value = "0"
            End If
            If Len(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")) = 0 Then
                ThisForm.GenerateEvent("Reprice")
            Else
                If CDec(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")) = 0 Then
                    'ThisForm.GenerateEvent "Reprice" - Issue 7798
                    ThisForm.GenerateEvent("CalculateNetPrice")
                Else
                    If ThisForm.PrimaryIDOCollection.IsCurrentObjectPropertyModified("PriceConv") Then
                        ThisForm.GenerateEvent("CalculateNetPrice")
                        'TRK 117336 - remove setting CalculatePrice = 0
                        ThisForm.Variables("OldUnitPrice").Value = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")
                        ThisForm.Variables("UnitPriceVar").Value = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PriceConv")
                    End If
                End If
            End If
        End Sub

        Sub ResetStat()
            ' force a value so that filtering will not get confused
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("Stat", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat"))
            ThisForm.PrimaryIDOCollection.NotifyDependentsToRefresh("Stat")
            If ThisForm.ParentFormName = "CustomerOrders" Then
                If ThisForm.ParentForm.PrimaryIDOCollection.GetCurrentObjectProperty("Type") = "B" Then
                    ThisForm.PostEvent("StdFormClose")
                End If
            End If
        End Sub

        Sub ResetCopiedvalues()
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("QtyReady", "0")
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("QtyRsvd", "0")
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("QtyPacked", "0")
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("QtyInvoiced", "0")
            ThisForm.PrimaryIDOCollection.SetCurrentObjectProperty("QtyShipped", "0")
            '    ThisForm.Variables("Action_copy") = 1
        End Sub

        Sub RefreshStat()
            ThisForm.PrimaryIDOCollection.NotifyDependentsToRefresh("Stat")
        End Sub

        Sub SetCreateSubJobsFlag()
            Dim vResponse As String


            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr") <> "" And
                 ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("RefType") = "J" Then

                vResponse = ThisForm.CallGlobalScript("MsgApp", "Clear", "Yes|No", "SuccessFailure",
                                                              "mQ=FunctPerformYesNo", "@sCreateSubJobs",
                                                              "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                ' convert to SQL concept of "true"
                If vResponse = "1" Then
                    vResponse = "0"
                Else
                    vResponse = "1"
                End If
                ThisForm.Variables("vResponse").Value = vResponse
                ThisForm.GenerateEvent("DefineVariable")
            End If

        End Sub

        Function ChkCustItemChanged() As Integer
            Dim oCache As IWSIDOCollection
            oCache = ThisForm.PrimaryIDOCollection

            ' CustItem was changed during update of an existing record - no need to continue
            If Not ThisForm.PrimaryIDOCollection.IsCurrentObjectNew And
               UCase(oCache.GetCurrentObjectProperty("CustItem")) <>
               UCase(ThisForm.Variables("PrevCustItem").Value) Then
                ChkCustItemChanged = 1
            Else ' CustItem did not change during update of an existing record - continue
                ChkCustItemChanged = 0
            End If
        End Function

        Sub ChkItemChanged()
            Dim oCache As IWSIDOCollection
            oCache = ThisForm.PrimaryIDOCollection

            If Not ThisForm.PrimaryIDOCollection.IsCurrentObjectNew And
               UCase(oCache.GetCurrentObjectProperty("Item")) <> UCase(ThisForm.Variables("CurrItem").Value) Then
                If UCase(oCache.GetCurrentObjectProperty("CustItem")) <> UCase(ThisForm.Variables("PrevCustItem").Value) Then
                    ThisForm.Variables("PrevCustItem").Value = oCache.GetCurrentObjectProperty("CustItem")
                End If

                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("PromotionCode", "")
            End If
        End Sub


        Sub SetCustItemValue()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("UbItemCustAdd") = "1" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("CustItem", ThisForm.Variables("tempCustItem").Value)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DueDate", ThisForm.Variables("tempDueDate").Value)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DerDueDate", ThisForm.Variables("tempDueDate").Value)
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("UbItemCustAdd") = "0" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item") = ThisForm.Variables("CurrItem").Value Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("CustItem", ThisForm.Variables("tempCustItem").Value)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DueDate", ThisForm.Variables("tempDueDate").Value)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DerDueDate", ThisForm.Variables("tempDueDate").Value)
            End If

            ThisForm.Variables("tempCustItem").Value = ""
        End Sub

        Sub FeatureGridToFeatStr()
            Dim oGroup As IWSIDOCollection
            Dim sFeatStr As String
            Dim i As Integer
            Dim iOffset As Integer
            Dim iLength As Integer
            Dim sOptCode As String
            Dim j As Integer
            Dim oMaterial As IWSIDOCollection
            Dim iLen As Integer
            Dim sFbomBlank As String

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItPlanFlag") <> "1" Then
                Exit Sub
            End If

            'Initializing Grids IDOCollection before loading current items feature/options
            If ThisForm.Components("FeatureGroupsGrid").IDOCollection.GetNumEntries > 0 Then
                ThisForm.Components("FeatureGroupsGrid").IDOCollection.ClearEntries()
            End If

            oGroup = ThisForm.Components("FGFeatureGridCol").IDOCollection

            sFeatStr = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr")

            sFbomBlank = ThisForm.Variables("InvparmsFbomBlank").Value
            If sFbomBlank = "" Then
                ThisForm.GenerateEvent("GetFbomBlank")
                sFbomBlank = ThisForm.Variables("InvparmsFbomBlank").Value
            End If


            For i = 0 To oGroup.GetNumEntries - 1
                iOffset = CInt(oGroup.GetObjectProperty("FeatureCodeOffset", i))
                iLength = CInt(oGroup.GetObjectProperty("FeatureCodeLength", i))

                iLen = (iOffset + iLength - 1) - Len(sFeatStr)
                If iLen > 0 Then
                    sFeatStr = sFeatStr + New String(Convert.ToChar(sFbomBlank), iLen)
                End If

                oGroup.MoveCurrentIndexAndRefresh(i, False)

                oMaterial = oGroup.GetSubCollection("SLJobmatls", -1)
                For j = 0 To oMaterial.GetNumEntries - 1
                    If oMaterial.GetObjectProperty("DerSelected", j) = "1" Then
                        sOptCode = ""
                        sOptCode = UCase(oMaterial.GetObjectProperty("OptCode", j))
                        Mid(sFeatStr, iOffset, iLength) = sOptCode

                    End If
                Next j
            Next i

            If Trim(Replace(sFeatStr, ",", "")) = "" Or Trim(Replace(sFeatStr, ",", "")) = "." Then
                sFeatStr = "." 'sFbomBlank
            End If

            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("FeatStr", sFeatStr)

            If oGroup.GetNumEntries > 0 Then
                oGroup.MoveCurrentIndexAndRefresh(0, False)
            End If

            oGroup = Nothing
            oMaterial = Nothing

        End Sub






        Sub AddItemLocation()
            If ThisForm.Variables("TrnLocQuestionAsked").Value = "1" Then
                ThisForm.CallGlobalScript("Ask",
                                          "PromptMsg",
                                          ThisForm.Variables("PromptButtons").Value,
                                          "ReturnUserActionVar", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

                If ThisForm.Variables("ReturnUserActionVar").Value = CStr(vbYes) Then
                    ThisForm.GenerateEvent("AddTrnLoc")
                End If
                ThisForm.GenerateEvent("CreateTransferOrderAgain")
            End If
        End Sub

        Sub SetOrignalGridQty()
            Dim oMaterial As IWSIDOCollection
            'Dim oComp As Object
            Dim i As Integer
            oMaterial = ThisForm.Components("FeatureMaterialsGrid").IDOCollection
            'oComp = ThisForm.Components("FeatureMaterialsViewGrid")
            ThisForm.Variables("CurrMatlQtyConv").Value = oMaterial.GetCurrentObjectProperty("MatlQtyConv")

            For i = 0 To oMaterial.GetNumEntries - 1
                If ThisForm.Components("FeatureMaterialsViewGrid").GetGridValueByColumnName(ThisForm.Components("FeatureMaterialsViewGrid").GetGridCurrentRow, "FMItemViewGridCol") = oMaterial.GetObjectProperty("Item", i) Then
                    oMaterial.MoveCurrentIndexAndRefresh(i, False)
                    If oMaterial.GetObjectProperty("UbQtyConv", i) = "" Then
                        oMaterial.SetObjectProperty("UbQtyConv",
                                                    i,
                                                    oMaterial.GetObjectProperty("MatlQtyConv", i))
                    End If
                    If CDec(oMaterial.GetObjectProperty("UbQtyConv", i)) = ThisForm.Components("FMMatlQtyViewGridCol").GetValueOfDecimal(0) Then
                        oMaterial.SetObjectProperty("UbQtyConv", i, "")
                    End If
                    oMaterial.SetObjectProperty("MatlQtyConv",
                                                i,
                                                ThisForm.Components("FeatureMaterialsViewGrid").GetGridValueByColumnName(ThisForm.Components("FeatureMaterialsViewGrid").GetGridCurrentRow, "FMMatlQtyViewGridCol"))
                    oMaterial.MoveCurrentIndexAndRefresh(i, False)
                End If
            Next
            ThisForm.Components("FeatureMaterialsViewGrid") = Nothing
            oMaterial = Nothing
        End Sub


        Sub ShowMessageQtyRsvd()
            Dim dPrevQtyOrderedConv As Decimal
            Dim dQtyOrderedConv As Decimal
            Dim dQtyRsvd As Decimal
            Dim iIndex As Integer
            If IsNumeric(ThisForm.Variables("PrevQtyOrderedConv")) Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dPrevQtyOrderedConv = ThisForm.PrimaryIDOCollection(iIndex)("PrevQtyOrderedConv").GetValueOfDecimal(0)

            Else
                dPrevQtyOrderedConv = CDec(0)
            End If

            If IsNumeric(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyOrderedConv")) Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQtyOrderedConv = ThisForm.PrimaryIDOCollection(iIndex)("QtyOrderedConv").GetValueOfDecimal(0)
            Else
                dQtyOrderedConv = CDec(0)
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("QtyRsvd") <> "" Then
                iIndex = ThisForm.CurrentIDOCollection.GetCurrentObjectIndex
                dQtyRsvd = ThisForm.PrimaryIDOCollection(iIndex)("QtyRsvd").GetValueOfDecimal(0)

            Else
                dQtyRsvd = 0
            End If

            If dQtyRsvd <> 0 And dQtyRsvd <> dQtyOrderedConv Then
                ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                "mI=IsCompare<>", "@sQtyOrdered", "@sQtyReserved", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End Sub

        Sub EnableDisableConfigButton()

            Dim bCfgEnabled As Boolean

            ' SyteLine Configurator
            bCfgEnabled = Application.Variables("Avail_Cfg").Value = "1" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItOrderConfigurable") = "1"

            If Not bCfgEnabled Then
                If Not Trim(Replace(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("FeatStr"), ",", "")) = "" And
                   ThisForm.Components("ConfigPushButton").Caption = Application.GetStringValue("sCreate&Item") Then
                    ThisForm.Variables("ConfigPushEnabled").SetValue("1")
                Else
                    ThisForm.Variables("ConfigPushEnabled").SetValue("0")
                End If
            End If
        End Sub

        Function ChkCOStatus() As Integer
            Dim sCOStat As String
            Dim sCustOrderStatus As String

            Dim sCoNum As String

            ChkCOStatus = 0
            sCOStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoStat")

            If sCOStat = "" Then
                ChkCOStatus = -1
                Exit Function
            End If

            sCustOrderStatus = CStr(Switch(sCOStat = "P", "Planned", sCOStat = "O", "Ordered" _
            , sCOStat = "C", "Complete", sCOStat = "S", "Stopped", sCOStat = "H", "History"))

            sCoNum = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")

            If sCOStat = "C" Or sCOStat = "H" Then
                ThisForm.PrimaryIDOCollection.SaveEnabled = False
                ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                        "mI=IsCompare1", "@sStatus", sCustOrderStatus, "@sCO", "@sOrder", sCoNum, "", "", "", "", "", "", "", "", "", "", "")
                ThisForm.SetFocus("CoNumEdit")
                ChkCOStatus = -1
                Exit Function
            Else
                ThisForm.PrimaryIDOCollection.SaveEnabled = True

            End If
        End Function
        Function ChkCOStatusForSave() As Integer
            Dim sCOStat As String

            ChkCOStatusForSave = 0
            sCOStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoStat")
            If sCOStat = "" Then
                ChkCOStatusForSave = -1
                Exit Function
            End If

            If sCOStat = "C" Or sCOStat = "H" Then
                ThisForm.PrimaryIDOCollection.SaveEnabled = False
                ChkCOStatusForSave = -1
                Exit Function
            Else
                ThisForm.PrimaryIDOCollection.SaveEnabled = True
            End If
        End Function

        Sub EnablePrintKitComponents()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItKit") <> "1" Then
                ThisForm.PrimaryIDOCollection.SetObjectPropertyPlusModifyRefresh("PrintKitComponents",
                    ThisForm.PrimaryIDOCollection.GetCurrentObjectIndex, "0")
            End If
        End Sub

        Sub CustOrderLinesFormDefaults()
            If ThisForm.Variables("Action_copy").Value <> "1" Then
                ThisForm.GenerateEvent("CustOrderLinesFormDefaults")
            Else
                ThisForm.Components("RefNumEdit").Text = ""
                ThisForm.Components("RefNumGridCol").Text = ""
                ThisForm.Components("RefLineSufEdit").Text = "0"
                ThisForm.Components("RefLineSufGridCol").Text = "0"
                ThisForm.Components("RefReleaseEdit").Text = "0"
                ThisForm.Components("RefReleaseGridCol").Text = "0"
            End If
        End Sub

        Function ValidateReservableItem() As Integer
            Dim StrButtonPressed As String
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItReservable") = "0" Then
                ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                          "mQ=CmdInvalid2", "@sAutomaticReservation", "@sItem", "@sItem", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item"),
                          "@sReservable", "0", "", "", "", "", "", "", "", "", "", "")
                ValidateReservableItem = -1
            Else
                StrButtonPressed = ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                          "mQ=CmdPerform2NoYes", "@sAutomaticReservation", "@sCOLine/Release", "@sOrder", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum"),
                          "@sLine", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoLine"), "", "", "", "", "", "", "", "", "", "")
                If StrButtonPressed = "0" Then
                    ValidateReservableItem = 0
                Else
                    ValidateReservableItem = -1
                End If
            End If

        End Function

        Sub EnableDisableComp()
            'ProdConfTab,ConfigPushButton,FeatStrEdit,FeatStrGridCol
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item") <> ThisForm.Variables("ItemVar").Value Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Item", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item"))
                ThisForm.PrimaryIDOCollection.SetCurrentObjectModified(True)
                ThisForm.Variables("ProdConfTabEnabled").SetValue("0")
                ThisForm.Variables("ConfigPushEnabled").SetValue("0")
                ThisForm.Variables("FeatStrEnabled").SetValue("0")
            End If
        End Sub

        Sub SetPlanOnSave()
            If ThisForm.Variables("ApsParmApsmode").Value <> "T" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("UbPlanOnSave", "1")
            End If


        End Sub

        Sub SetPages()
            If ThisForm.Components(ThisForm.Components("Notebook").NotebookCurTab).Enabled = False Then
                ThisForm.Components("Notebook").NotebookCurTab = "Tab1"
            End If
        End Sub

        Sub ResetIncrPrice()
            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("UbIncrPrice", "0")
        End Sub

        Sub DefaultDropShipSeq()
            ThisForm.Components("CustSeqEdit").Text = "0"
            ThisForm.Components("CustSeqGridCol").Text = "0"
        End Sub


        Sub SetFeatQtysVar()
            Dim oGroup, oMaterial As IWSIDOCollection
            Dim i, j As Integer
            Dim FeatureQtys, Feature, Qtys As String

            FeatureQtys = ""
            Feature = ""
            Qtys = ""
            oGroup = ThisForm.Components("FGFeatureGridCol").IDOCollection

            For i = 0 To oGroup.GetNumEntries - 1
                Feature = oGroup.GetObjectProperty("Feature", i)
                oGroup.MoveCurrentIndexAndRefresh(i, False)
                oMaterial = oGroup.GetSubCollection("SLJobmatls", i)
                For j = 0 To oMaterial.GetNumEntries - 1
                    If oMaterial.GetObjectProperty("DerSelected", j) = "1" Then
                        If oMaterial.GetObjectProperty("UbQtyConv", j) <> "" _
                        And oMaterial.Item(j).Properties("UbQtyConv").GetValueOfDecimal(0) <> oMaterial.Item(j).Properties("MatlQtyConv").GetValueOfDecimal(0) Then
                            Dim Multiplier As Int64 = 100000000
                            Qtys = CStr(oMaterial.Item(j).Properties("MatlQtyConv").GetValueOfDecimal(0) * Multiplier)
                            FeatureQtys = FeatureQtys + Feature + "," + Qtys
                        End If
                        Exit For
                    End If
                Next j
                If Not i = oGroup.GetNumEntries - 1 Then
                    FeatureQtys = FeatureQtys + ","
                End If
            Next i

            ThisForm.Variables("FeatQtys").Value = FeatureQtys
            oGroup = Nothing
            oMaterial = Nothing
        End Sub

        Function CheckUserPermissions() As Integer
            Dim UserHasPermission As String

            UserHasPermission = ""
            CheckUserPermissions = 0
            If ThisForm.Variables("CreateFlag").Value = "1" Or ThisForm.Variables("CreateFlag1").Value = "1" Then
                ThisForm.Variables("vAction").Value = "Insert"
                Select Case ThisForm.Components("RefTypeEdit").ValueInternal
                    Case "I" 'Inventory
                        ThisForm.Variables("vFormName").Value = "Items"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                    Case "P" 'Purchase Orders
                        ThisForm.Variables("vFormName").Value = "PurchaseOrders"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                    Case "J" 'Job Orders
                        ThisForm.Variables("vFormName").Value = "JobOrders"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                    Case "R" 'Requistions
                        ThisForm.Variables("vFormName").Value = "PurchaseOrderRequisitions"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                    Case "T" 'Transfers
                        ThisForm.Variables("vFormName").Value = "TransferOrders"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                    Case "K" 'Projects
                        ThisForm.Variables("vFormName").Value = "Projects"
                        ThisForm.GenerateEvent("FormPermissionsCheck")
                        UserHasPermission = ThisForm.Variables("VCanDo").Value

                End Select
            End If

            If UserHasPermission = "0" Then
                ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure",
                         "mE=NoFunct", ThisForm.Variables("vAction").Value, ThisForm.Variables("vFormName").Value, "", "", "", "", "",
                          "", "", "", "", "", "", "", "", "")
                CheckUserPermissions = -1
                Exit Function
            End If

        End Function


        Sub ValidateWhseChanged()
            Dim bWhse As Boolean
            bWhse = (ThisForm.Variables("PrevWhse").Value <>
                             ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Whse"))
            If bWhse Then
                ThisForm.Variables("CalculatePrice").Value = "0"
                ThisForm.GenerateEvent("DoGeneratePriceCalculation")
            End If

        End Sub

        Sub ValidateShipSiteChanged()
            Dim bShipSite As Boolean
            bShipSite = (UCase(ThisForm.Variables("PrevSite").Value) <>
                             UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")))
            ThisForm.Variables("ChangeShipSite").Value = "0"
            If bShipSite Then
                ThisForm.Variables("CurrItem").Value = ""
                ThisForm.Variables("ValidateShipSiteChanged").Value = "1"
                ThisForm.GenerateEvent("ItemChanged")
                ThisForm.Variables("ValidateShipSiteChanged").Value = ""
            End If
            ThisForm.Variables("ChangeShipSite").Value = "1"
        End Sub

        Sub SetShipSite()
            If ThisForm.Variables("ChangeShipSite").Value = "1" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("ShipSite", ThisForm.Variables("vShipSite").Value)
            End If
        End Sub

        Sub ReturnFromConfig()
            If ThisForm.LastModalChildName = "Configuration" Then
                If Application.Variables("CpqWarning").Value = "1" Then
                    ThisForm.GenerateEvent("CpqWarning")
                    Application.Variables("CpqWarning").SetValue("")
                End If
                ThisForm.PrimaryIDOCollection.RefreshKeepCurIndex()
            End If
        End Sub

        Sub RefreshIfCrossSite()
            If Not String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite"),
                                 ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                 StringComparison.OrdinalIgnoreCase) Then
                ThisForm.PrimaryIDOCollection.RefreshCurrentObject()
            End If
        End Sub

        Sub SaveLineIfModified()
            If ThisForm.PrimaryIDOCollection.IsCurrentObjectModified Then
                ThisForm.PrimaryIDOCollection.SaveCurrent()
            End If
        End Sub

        Function ItemValidate() As Integer
            Dim iNonInventoryItem As Integer
            Dim bNonInvenrotyItem As Boolean
            If ThisForm.Variables("ItemExists").Value = "1" Then
                iNonInventoryItem = 0
            Else
                iNonInventoryItem = 1
            End If
            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DerNonInventoryItemFlag", iNonInventoryItem.ToString)
            bNonInvenrotyItem = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerNonInventoryItemFlag") = "1"
            If Not bNonInvenrotyItem Then
                Return ThisForm.GenerateEvent("ItemValidateInventory")
            End If
            Return 0
        End Function

        Sub ItemChangedInventoryItem()
            Dim bNonInvenrotyItem As Boolean
            bNonInvenrotyItem = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerNonInventoryItemFlag") = "1"
            If Not bNonInvenrotyItem Then
                ThisForm.GenerateEvent("ItemChangedInventoryItem")
            End If
        End Sub

        Sub RefreshDIFOT()
            Dim dShippedOverOrderedQtyTolerance As Decimal = Decimal.Zero
            Dim dShippedUnderOrderedQtyTolerance As Decimal = Decimal.Zero
            Dim iDaysShippedBeforeDueDateTolerance As Integer = 0
            Dim iDaysShippedAfterDueDateTolerance As Integer = 0
            Dim dDerQtyShippedOver As Decimal = Decimal.Zero
            Dim dDerQtyShippedUnder As Decimal = Decimal.Zero
            Dim iDerShippedBeforeDueDate As Integer = 0
            Dim iDerShippedAfterDueDate As Integer = 0

            Dim bShippedOverOrderedQtyTolerance As Boolean
            Dim bShippedUnderOrderedQtyTolerance As Boolean
            Dim bDaysShippedBeforeDueDateTolerance As Boolean
            Dim bDaysShippedAfterDueDateTolerance As Boolean

            'Get Read Only property
            If Trim(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerQtyShippedOver")) <> "" Then
                dDerQtyShippedOver = CDec(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerQtyShippedOver"))
            End If

            If Trim(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerQtyShippedUnder")) <> "" Then
                dDerQtyShippedUnder = CDec(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerQtyShippedUnder"))
            End If

            If Trim(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerShippedBeforeDueDate")) <> "" Then
                iDerShippedBeforeDueDate = CInt(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerShippedBeforeDueDate"))
            End If

            If Trim(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerShippedAfterDueDate")) <> "" Then
                iDerShippedAfterDueDate = CInt(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerShippedAfterDueDate"))
            End If

            If Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShippedOverOrderedQtyTolerance")) Then
                dShippedOverOrderedQtyTolerance = CDec(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShippedOverOrderedQtyTolerance"))
            Else
                bShippedOverOrderedQtyTolerance = True
            End If

            If Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShippedUnderOrderedQtyTolerance")) Then
                dShippedUnderOrderedQtyTolerance = CDec(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShippedUnderOrderedQtyTolerance"))
            Else
                bShippedUnderOrderedQtyTolerance = True
            End If

            If Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DaysShippedBeforeDueDateTolerance")) Then
                iDaysShippedBeforeDueDateTolerance = CInt(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DaysShippedBeforeDueDateTolerance"))
            Else
                bDaysShippedBeforeDueDateTolerance = True
            End If

            If Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DaysShippedAfterDueDateTolerance")) Then
                iDaysShippedAfterDueDateTolerance = CInt(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DaysShippedAfterDueDateTolerance"))
            Else
                bDaysShippedAfterDueDateTolerance = True
            End If

            If Not bShippedOverOrderedQtyTolerance Then
                bShippedOverOrderedQtyTolerance = dDerQtyShippedOver <= dShippedOverOrderedQtyTolerance
            End If

            If Not bShippedUnderOrderedQtyTolerance Then
                bShippedUnderOrderedQtyTolerance = dDerQtyShippedUnder <= dShippedUnderOrderedQtyTolerance
            End If

            If Not bDaysShippedBeforeDueDateTolerance Then
                bDaysShippedBeforeDueDateTolerance = iDerShippedBeforeDueDate <= iDaysShippedBeforeDueDateTolerance
            End If

            If Not bDaysShippedAfterDueDateTolerance Then
                bDaysShippedAfterDueDateTolerance = iDerShippedAfterDueDate <= iDaysShippedAfterDueDateTolerance
            End If

            If bShippedOverOrderedQtyTolerance And bShippedUnderOrderedQtyTolerance Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerInFull"), "0")
            Else
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerInFull"), "1")
            End If

            If bDaysShippedBeforeDueDateTolerance And bDaysShippedAfterDueDateTolerance Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerOnTime"), "0")
            Else
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerOnTime"), "1")
            End If

            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerInFull") = "0" And
                ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerOnTime") = "0" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerDIFOT"), "0")
            Else
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh(("DerDIFOT"), "1")
            End If

        End Sub

        Sub RefreshDIFOTOnStatus()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "F" _
                Or ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat") = "C" Then

                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyModified("ShippedOverOrderedQtyTolerance", False)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyModified("ShippedUnderOrderedQtyTolerance", False)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyModified("DaysShippedBeforeDueDateTolerance", False)
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyModified("DaysShippedAfterDueDateTolerance", False)
            End If

        End Sub

        Sub RefreshItem()
            If (Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item"))) And
                UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item")) <> UCase(ThisForm.Variables("CurrItem").Value) Then
                ThisForm.Components("ItemEdit").ValidateData(True)
                ThisForm.GenerateEvent("ItemChanged")
            End If
        End Sub

        Sub SetShipSiteComponentClass()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoConsignment") = "1" Then
                ThisForm.Variables("ShipSiteIDOVar").Value = "SLWhseAlls"
                ThisForm.Variables("ShipSitePropertiesVar").Value = "SiteRef,SiteSiteName, Whse, Name"
                ThisForm.Variables("ShipSiteFilterVar").Value = " CustNum = FP(CoCustNum) AND CustSeq = FP(CoCustSeq) " +
                    "AND ConsignmentType = 'C' "
                ThisForm.Variables("ShipSiteValidatorVar").Value = "ShipSiteForConsigned(CoCustNum,CoCustSeq,Whse,),SiteForShip(CoOrigSite, ShipSite, custaddr_mst, 1),GenerateEventNoMsg(ItemValidate)"
            Else
                ThisForm.Variables("ShipSiteIDOVar").Value = "SLSites"
                ThisForm.Variables("ShipSitePropertiesVar").Value = "Site, SiteName"
                ThisForm.Variables("ShipSiteFilterVar").Value = " IntIsExternal <> 1 or IntranetName = NULL)"
                ThisForm.Variables("ShipSiteValidatorVar").Value = "Site,SiteForShip(CoOrigSite, ShipSite, custaddr_mst, 1),GenerateEventNoMsg(ItemValidate)"
            End If
        End Sub

        Sub SetManufacturerMenu()
            If String.Equals(Application.Variables("Parm_Site").Value, ThisForm.CurrentIDOCollection.GetCurrentObjectProperty("ShipSite"), StringComparison.OrdinalIgnoreCase) Or
                ThisForm.CurrentIDOCollection.GetCurrentObjectProperty("ShipSite") = "" Then
                ThisForm.Components("ManufacturerIdEdit").MenuName = "StdDetailsAddFind"
                ThisForm.Components("ManufacturerIdGridCol").MenuName = "StdDetailsAddFind"
                ThisForm.Components("ManufacturerItemEdit").MenuName = "StdDetailsAddFind"
                ThisForm.Components("ManufacturerItemGridCol").MenuName = "StdDetailsAddFind"
            Else
                ThisForm.Components("ManufacturerIdEdit").MenuName = "StdDefault"
                ThisForm.Components("ManufacturerIdGridCol").MenuName = "StdDefault"
                ThisForm.Components("ManufacturerItemEdit").MenuName = "StdDefault"
                ThisForm.Components("ManufacturerItemGridCol").MenuName = "StdDefault"
            End If
        End Sub

        Sub ItemValidateAsk()
            Dim bFranceCP As Boolean
            Dim bCSIB_97770 As Boolean
            Dim allowNonItemVar As Boolean

            bCSIB_97770 = Application.Variables("IsCSIB_97770ActiveVar").GetValueOfInteger(0) = 1
            bFranceCP = Application.Variables("Avail_FR").GetValueOfInteger(0) = 1
            allowNonItemVar = ThisForm.Variables("AllowNonItemVar").GetValueOfInteger(0) = 1

            If Not (bCSIB_97770 And bFranceCP And Not allowNonItemVar) Then
                'Do not pop up message if item and shipsite are not changed.
                If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item") <> ThisForm.Variables("CurrItem").Value _
                    Or (Not String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                 ThisForm.Variables("PrevSite").Value, StringComparison.OrdinalIgnoreCase)) Then
                    If String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                 Application.Variables("Parm_Site").Value, StringComparison.OrdinalIgnoreCase) Then
                        ThisForm.CallGlobalScript("Ask", "MessageVar", "Cancel|OK", "PromptOpenResponse", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                        ThisForm.CallGlobalScript("EvalResponseAndVariableGenerateEventWithNoError", "PromptOpenResponse", "IsOpenNonInvForm", "1", "LaunchFormNonInvItem", "ItemEdit", "FALSE", "P", "", "", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        ThisForm.CallGlobalScript("Ask", "MessageVar", "OK", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If
            End If
        End Sub

        Function GetCoNonInvAcct() As Integer

            Dim bIsNonInventoryItem As Boolean
            Dim bIsPostjour As Boolean

            'We only want to retrieve G/L account if this is NonInventory item, or PostJour is False
            GetCoNonInvAcct = 0
            bIsNonInventoryItem = ThisForm.Variables("ItemExists").Value <> "1"
            bIsPostjour = ThisForm.Variables("PostJourVar").Value = "1"
            If bIsNonInventoryItem = False And bIsPostjour = True Then
                Return GetCoNonInvAcct
            End If

            'Do not retrieve G/L account if item and shipsite are not changed.
            GetCoNonInvAcct = 0
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Item") <> ThisForm.Variables("CurrItem").Value _
            Or Not String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                 ThisForm.Variables("PrevSite").Value, StringComparison.OrdinalIgnoreCase) Then
                GetCoNonInvAcct = ThisForm.GenerateEvent("GetCoNonInvAcct")
            End If
        End Function

        Function CheckIfWebClient() As Integer
            If Application.Platform = "WEB" Then
                Call ThisForm.CallGlobalScript("MsgApp", "Clear", "Ok", "", "mFormNotPermittedOnWebClient", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                Return 1
            End If
            Return 0
        End Function

        Sub EnableDisableNewInit()
            'Prevent the AutoInsert row from being created when opening linked from a CO and new rows should not be added.
            Dim sCOStat As String

            sCOStat = ThisForm.Variables("CoStatVar").Value

            If sCOStat = "C" And ThisForm.ParentFormName <> "" Then
                ThisForm.PrimaryIDOCollection.NewEnabled = False
            End If

        End Sub

        Sub GetTransferDefaults()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoConsignment") = "0" Then
                ThisForm.GenerateEvent("GetTransferDefaults")
            End If
        End Sub

        Sub GetPostJourValue()
            Dim oParmsCache As LoadCollectionResponseData
            Dim sPostJour As String

            'get PostJour
            oParmsCache = IDOClient.LoadCollection("SL.SLParms", New PropertyList("PostJour"), "", "", 0)
            sPostJour = (oParmsCache(0, "PostJour").ToString)
            ThisForm.Variables("PostJourVar").Value = sPostJour
        End Sub

        Sub InitialiseItem()
            If ThisForm.Variables("ValidateShipSiteChanged").Value <> "1" _
                And ThisForm.Variables("PrevCustItem").Value = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CustItem") Then
                If ThisForm.Variables("ItemValueBeforeManufacturerItemChanged").Value <> ThisForm.Components("ItemEdit").Value Then
                    ThisForm.GenerateEvent("InitialiseItemInfo")
                End If
            End If
        End Sub

        Function SetItemValueBeforeManufacturerItemChanged() As Integer
            ThisForm.Variables("ItemValueBeforeManufacturerItemChanged").Value = ThisForm.Components("ItemEdit").Value
            SetItemValueBeforeManufacturerItemChanged = 0
        End Function

        Sub EnableDemandingSite()
            'ItemEdit,ItemGridCol
            'PriceConvEdit,PriceConvGridCol
            'DiscEdit,DiscGridCol
            'CustItemEdit,CustItemGridCol
            'QtyOrderedConvEdit,QtyOrderedConvGridCol
            'UMEdit,UMGridCol
            'DueDateEdit,DueDateGridCol
            'CustSeqEdit,CustSeqGridCol
            'CustNumEdit,CustNumGridCol
            'DropShipAddressEdit
            If ThisForm.Components("DemandingSiteEdit").Value <> "" Then
                ThisForm.Variables("ItemEnabled").SetValue("0")
                ThisForm.Variables("PriceConvEnabled").SetValue("0")
                ThisForm.Variables("DiscEnabled").SetValue("0")
                ThisForm.Variables("CustItemEnabled").SetValue("0")
                ThisForm.Variables("QtyOrderedConvEnabled").SetValue("0")
                ThisForm.Variables("UMEnabled").SetValue("0")
                ThisForm.Variables("DueDateEnabled").SetValue("0")
                ThisForm.Variables("CustSeqEnabled").SetValue("0")
                ThisForm.Variables("CustNumEnabled").SetValue("0")
                ThisForm.Variables("DropShipAddressEnabled").SetValue("0")
                ThisForm.PrimaryIDOCollection.DeleteEnabled = False
                ThisForm.PrimaryIDOCollection.NewEnabled = False
            Else
                'ThisForm.PrimaryIDOCollection.DeleteEnabled = True
                'ThisForm.PrimaryIDOCollection.NewEnabled = True
            End If
        End Sub

        Sub WarnMaxCoDisc()
            If ThisForm.Variables("CoNumList").Value <> "" Then
                ThisForm.CallGlobalScript("MsgApp", "Clear", "NoPrompt", "SuccessFailure", "mI=Changed",
                                          "@sOrderDisc", "@sCustomerOrder", "", "", "",
                                          "", "", "", "", "", "", "", "", "", "", "")
                ThisForm.CallGlobalScript("MsgApp", "NoClear", "NoPrompt", "SuccessFailure", "mW=MayNeedToRecalculate",
                                          "@sTotalPrice", "@sOrderDisc", "", "", "",
                                          "", "", "", "", "", "", "", "", "", "", "")
                ThisForm.CallGlobalScript("MsgApp", "NoClear", "Prompt", "SuccessFailure", "mE=IsCompare",
                                          "@sCustomerOrder", ThisForm.Variables("CoNumList").Value, "", "", "",
                                          "", "", "", "", "", "", "", "", "", "", "")
            End If
        End Sub

        Sub RedefaultWhse()
            If Not String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite"),
                                 ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                 StringComparison.OrdinalIgnoreCase) And
                ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoConsignment") = "0" Then
                ThisForm.GenerateEvent("GetTransferDefaults")
            End If
        End Sub

        Sub GetNonInvItemInfo()
            Dim bNonInvenrotyItem As Boolean = False
            bNonInvenrotyItem = ThisForm.Variables("ItemExists").GetValueOfString("") = "0"
            If bNonInvenrotyItem AndAlso
                 String.Equals(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"),
                                Application.Variables("Parm_Site").Value, StringComparison.OrdinalIgnoreCase) Then
                ThisForm.GenerateEvent("GetNonInvItemInfo")
                CalculateNetPrice()
                ThisForm.Variables("RightClickItemForm").Value = "Non-InventoryItems"
                ThisForm.Variables("RightClickItemQueryForm").Value = "Non-InventoryItemsQuery"
            Else
                ThisForm.Variables("RightClickItemForm").Value = "Items"
                ThisForm.Variables("RightClickItemQueryForm").Value = "ItemsQuery"
            End If
        End Sub
        Sub SetRightClickVariable()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerItExist") = "0" _
                AndAlso ThisForm.PrimaryIDOCollection.CurrentItem.Properties("Item").GetValueOfString("").Length > 0 Then
                ThisForm.Variables("RightClickItemForm").Value = "Non-InventoryItems"
                ThisForm.Variables("RightClickItemQueryForm").Value = "Non-InventoryItemsQuery"
            Else
                ThisForm.Variables("RightClickItemForm").Value = "Items"
                ThisForm.Variables("RightClickItemQueryForm").Value = "ItemsQuery"
            End If
        End Sub

        Sub MakeSalesDiscToZero()
            If ThisForm.PrimaryIDOCollection.CurrentItem("PromotionCode").GetValueOfString("") <> "" Then
                ThisForm.PrimaryIDOCollection.CurrentItem("Disc").SetValuePlusModifyRefresh(0)
            End If
        End Sub

        Function ValidateSalesDiscWithPromotionCode() As Integer
            ValidateSalesDiscWithPromotionCode = CInt(ThisForm.PrimaryIDOCollection.CurrentItem("PromotionCode").GetValueOfString("") <> "" _
                   AndAlso ThisForm.PrimaryIDOCollection.CurrentItem("Disc").GetValueOfDecimal(0) > 0)
        End Function

        Sub CheckConfigID()
            Dim i As Integer
            Dim currentIndex As Integer

            If Not ThisForm.Variables("vCopyFlag").Value = "1" Then
                Exit Sub
            End If

            currentIndex = ThisForm.PrimaryIDOCollection.GetCurrentObjectIndex()

            For i = 0 To ThisForm.PrimaryIDOCollection.GetNumEntries - 1
                If ThisForm.PrimaryIDOCollection.Items(i).Item("UbSourceCoNum").GetValueOfString("") <> "" And
                        ThisForm.PrimaryIDOCollection.Items(i).Item("UbSourceCoLine").GetValueOfString("") <> "" And
                        ThisForm.PrimaryIDOCollection.Items(i).Item("UbSourceConfigID").GetValueOfString("") <> "" Then
                    ThisForm.PrimaryIDOCollection.SetCurrentObject(i)
                    ThisForm.GenerateEvent("CopyConfig")
                    ThisForm.PrimaryIDOCollection.RefreshCurrentObject()
                End If
            Next

            ThisForm.PrimaryIDOCollection.SetCurrentObject(currentIndex)
            ThisForm.Variables("vCopyFlag").Value = "0"

        End Sub

        Function PromotionCodeChanged() As Integer
            PromotionCodeChanged = 0
            If ThisForm.Variables("CurPromotionCodeVar").GetValueOfString("") = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PromotionCode") Then
                PromotionCodeChanged = -1
            End If
        End Function

        Sub RefreshPromotionCodes()
            ThisForm.Components("PromotionCodeEdit").InvalidateList()
            ThisForm.Components("PromotionCodeGridCol").InvalidateList()
        End Sub

        Sub RunReprice()
            If Application.Variables("Avail_AU").Value = "1" Then
                ThisForm.GenerateEvent("RunRepriceForm")
            Else
                ThisForm.GenerateEvent("Reprice")
            End If
        End Sub

        Sub SetUnitPricePlusIncrPrice()
            Dim dIncrPrice As Decimal

            dIncrPrice = ThisForm.PrimaryIDOCollection.CurrentItem.Properties("UbIncrPrice").GetValueOfDecimal(0)

            With ThisForm.PrimaryIDOCollection.CurrentItem
                .Properties("PriceConv").SetValuePlusModifyRefresh(.Properties("PriceConv").GetValueOfDecimal(0) + dIncrPrice)
            End With
        End Sub

        ' Defaults that do not do anything are set as validated on new record.
        Sub SetDefaultsValidated()

            ThisForm.Components("ItemEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("ItemGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("StatEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("StatGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CustSeqEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CustSeqGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("WhseEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("WhseGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("QtyOrderedConvEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("QtyOrderedConvGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PromotionCodeEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PromotionCodeGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PriceConvEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PriceConvGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("RefTypeEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("RefTypeGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("ShipSiteEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("ShipSiteGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CommCodeEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CommCodeGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("UnitWeightEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("UnitWeightGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("TaxCode1Edit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("TaxCode1GridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("TaxCode2Edit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("TaxCode2GridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("OriginEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("OriginGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PricecodeEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("PricecodeGridCol").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CustItemEdit").SetModifiedSinceLoadOrValidation(False)
            ThisForm.Components("CustItemGridCol").SetModifiedSinceLoadOrValidation(False)

            ThisForm.Components("CoCustNumEdit").SetModifiedSinceLoadOrValidation(False)


        End Sub

        Sub StdObjectNewCompleted_BusinessLogic()
            ValidateCoNumEdit()
            CustOrderLinesFormDefaults()
            ValidateShipSiteEdit()
            SetOrderStatus()
            ApplyMask()
            EnableXrefButton()
            EnableFields()
            SetPlanOnSave()
            ChkCOStatus()
            ThisForm.GenerateEvent("GetDIFOTPolicy")
            SetShipSiteComponentClass()
        End Sub

        Sub StdObjectSelectCurrentCompleted_BusinessLogic()
            Dim bThisIsANew As Boolean
            bThisIsANew = ThisForm.CurrentIDOCollection.IsCurrentObjectNewAndUnmodified

            ThisForm.GenerateEvent("CreditOrderHoldDisplayCheck")
            ThisForm.GenerateEvent("DIFOTHoldDisplayCheck")
            If bThisIsANew = False Then  'If this is new, don't need this since done by StdObjectNewCompleted
                EnableXrefButton()
            End If
            ThisForm.GenerateEvent("ConfigSetup")
            RefreshStat()
            SetPages()
            ThisForm.GenerateEvent("GetCurrDecimalPlacesSp")
            RefreshDIFOTOnStatus()
            If bThisIsANew = False Then 'If this is new, don't need this since done by StdObjectNewCompleted
                EnableFields()
                SetShipSiteComponentClass()
            End If

            ThisForm.GenerateEvent("ReSetNonInvUnitCodeAccess")
            SetManufacturerMenu()
            SetRightClickVariable()
            EnableSourceTab()
            EnableDisableDelete()
        End Sub

        Private Sub EnableDisableDelete()
            Dim sStat As String
            Dim bPlanned As Boolean
            Dim bIsNew As Boolean
            Dim bOrdered As Boolean
            Dim bEnable As Boolean
            Dim bOrigSite As Boolean

            bIsNew = ThisForm.PrimaryIDOCollection.IsCurrentObjectNew
            sStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat")
            bPlanned = sStat = "P" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerOldStat") <> "C"
            bOrdered = sStat = "O" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerOldStat") <> "C"
            bOrigSite = (UCase(Application.Variables("Parm_Site").Value) =
                         UCase(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoOrigSite")))

            If (bPlanned Or bOrdered Or bIsNew) And bOrigSite Then
                bEnable = (ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("LastExternalShipDocID") = "")
                ThisForm.PrimaryIDOCollection.DeleteEnabled = bEnable
                'DueDateEdit,DueDateGridCol,AUDueDateEdit,AUDueDateGridCol
                If bEnable and ThisForm.PrimaryIDOCollection.CurrentItem.Properties("DemandingSite").GetValueOfString("") = "" Then
                    ThisForm.Variables("DueDateEnabled").SetValue("1")
                    ThisForm.Variables("AUDueDateEnabled").SetValue("1")
                Else
                    ThisForm.Variables("DueDateEnabled").SetValue("0")
                    ThisForm.Variables("AUDueDateEnabled").SetValue("0")
                End If
            End If
        End Sub

        Sub DueDate()
            If Application.Variables("Avail_AU").Value = "1" Then
                ThisForm.GenerateEvent("DataChangeOnDueDate")
            End If
        End Sub

        Sub CalcNewPrice()
            If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("AUPriceBy") = "D" Then
                ThisForm.CallGlobalScript("MsgApp", "Clear", "NoPrompt", "SuccessFailure",
                        "mI=ImpactsCosts", "@sDueDate", "@sPrice", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                If ThisForm.CallGlobalScript("MsgApp", "NoClear", "Prompt", "SuccessFailure",
                        "mQ=CostRecalcYesNo", "@sPrice", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "") = "0" Then
                    ThisForm.GenerateEvent("CalculateCoitemPrice")
                End If
            End If
        End Sub

        Sub CalculateCoitemPrice()
            If Application.Variables("Avail_AU").Value = "1" Then
                ThisForm.GenerateEvent("CalculatePriceOnAU")
            Else
                ThisForm.GenerateEvent("CalculatePrice")
            End If
        End Sub

        Function Validate_ATPCTP() As Integer
            Dim sPlanningMode As String
            Dim sStat As String
            Dim sCalcATPCTP As String
            Dim sParmReqSrc As String
            Dim bNonInventoryItem As Boolean
            Dim sPromiseDate As String

            Dim oParmsCache As LoadCollectionResponseData
            Dim oApsParmAll As LoadCollectionResponseData

            Dim oCOLineCollection As IWSIDOCollection
            oCOLineCollection = ThisForm.Components("FormCollectionGrid").IDOCollection

            sCalcATPCTP = "0"

            If Not String.IsNullOrEmpty(ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite")) Then
                Dim strFilter As String = String.Format("SiteRef = '{0}'", ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ShipSite"))
                oApsParmAll = IDOClient.LoadCollection("SLApsParmAlls", "CalcAtpCtpForAllCoLines", strFilter, "", -1)
                If Not oApsParmAll Is Nothing And oApsParmAll.Items.Count > 0 Then
                    sCalcATPCTP = oApsParmAll.Items(0).PropertyValues(0).Value
                End If
            Else
                oParmsCache = IDOClient.LoadCollection("SL.SLApsParms", New PropertyList("CalculateATPCTPForAllCOLines"), "", "", 0)
                sCalcATPCTP = (oParmsCache(0, "CalculateATPCTPForAllCOLines").ToString)
            End If

            sPlanningMode = ThisForm.Variables("ApsParmApsmode").Value
            sParmReqSrc = ThisForm.Variables("MrpParmReqSrc").Value
            bNonInventoryItem = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerNonInventoryItemFlag") = "1"

            sStat = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("Stat")
            sPromiseDate = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("PromiseDate")

            If sPlanningMode <> "T" And
               sParmReqSrc <> "F" And
               (sStat = "O" Or sStat = "P") And
               sCalcATPCTP = "1" And
               sPromiseDate = "" And
               Not bNonInventoryItem And
               Not ThisForm.PrimaryIDOCollection.IsCurrentObjectDeleted() Then
                If oCOLineCollection.IsCollectionModified = True Then
                    ThisForm.GenerateEvent("ShowCTPResults")
                    Return -1
                Else
                    Return 0
                End If
            End If
        End Function

        Sub GetCalcATPCTPValue()
            Dim oParmsCache As LoadCollectionResponseData
            Dim sCalcATPCTP As String

            'get CalculateATPCTPForAllCOLines
            oParmsCache = IDOClient.LoadCollection("SL.SLApsParms", New PropertyList("CalculateATPCTPForAllCOLines"), "", "", 0)
            sCalcATPCTP = (oParmsCache(0, "CalculateATPCTPForAllCOLines").ToString)
            ThisForm.Variables("Calculate_ATPCTP").Value = sCalcATPCTP
            ThisForm.Components("ApsModeCheck").Value = sCalcATPCTP
        End Sub

        Sub ChkCTPResultsReturnedValue()
            If (ThisForm.LastModalChildName = "CTPResults") Then
                If (ThisForm.LastModalChildEndedOk) Then
                    Dim oCOLineCollection As IWSIDOCollection
                    oCOLineCollection = ThisForm.Components("FormCollectionGrid").IDOCollection

                    If oCOLineCollection.IsCollectionModified = True Then
                        Application.ShowMessage(Application.GetStringValue("sPendingChangesPrompt"))
                    End If
                Else
                    ThisForm.CallGlobalScript("MsgApp", "Clear", "OK", "SuccessFailure",
                       "mI=IsCancelled", "@sCheckIn", "", "", "", "", "", "", "", "",
                       "", "", "", "", "", "", "")
                End If
            End If
        End Sub

        Function SaveAtEnabledGetATPCTP() As Integer
            Dim oCOLineCollection As IWSIDOCollection
            oCOLineCollection = ThisForm.PrimaryIDOCollection

            GetCalcATPCTPValue()

            Dim sCalcATPCTP As String
            sCalcATPCTP = ThisForm.Variables("Calculate_ATPCTP").Value

            If sCalcATPCTP = "1" Then
                If oCOLineCollection.IsCurrentObjectModified() = True Or oCOLineCollection.IsCurrentObjectNew() = True Then
                    Application.ShowMessage(Application.GetStringValue("sSaveAtEnableGetATPCTPPrompt"), vbOKOnly)
                    Return -1
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Function

        Function SelectAtEnabledGetATPCTP() As Integer
            Dim oCOLineCollection As IWSIDOCollection
            oCOLineCollection = ThisForm.PrimaryIDOCollection

            GetCalcATPCTPValue()

            Dim sCalcATPCTP As String
            sCalcATPCTP = ThisForm.Variables("Calculate_ATPCTP").Value

            If sCalcATPCTP = "1" Then
                If oCOLineCollection.IsCurrentObjectModified() = True Or oCOLineCollection.IsCurrentObjectNew() = True Then
                    Application.ShowMessage(Application.GetStringValue("sSelectAtEnableGetATPCTPPrompt"), vbOKOnly)
                    Return -1
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Function

        Sub RefreshNonInvAcct()
            'NonInvAcctUnit1Edit,NonInvAcctUnit1GridCol
            'NonInvAcctUnit2Edit,NonInvAcctUnit2GridCol
            'NonInvAcctUnit3Edit,NonInvAcctUnit3GridCol
            'NonInvAcctUnit4Edit,NonInvAcctUnit4GridCol
            If Not ThisForm.Components("NonInvAcctEdit").Enabled Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcct", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcctUnit1", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcctUnit2", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcctUnit3", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcctUnit4", "")
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("NonInvAcctDesc", "")
                ThisForm.Variables("EnableNonInvAcctCompsVar").SetValue("0")
            End If
        End Sub

        Sub SetPrimaryIDOAsCurrent()
            If ThisForm.CurrentIDOCollection.GetDisplayedObjectName() <> ThisForm.PrimaryIDOCollection.GetDisplayedObjectName() Then
                ThisForm.CurrentIDOCollection = ThisForm.PrimaryIDOCollection
            End If
        End Sub

        Sub SetFormInitials()
            If ThisForm.Variables("parm_DisplayTransactionNumInTitleBar").GetValueOfByte(0) = 1 Then
                Dim maxStringLength As Integer = 501
                Dim captionChars(maxStringLength) As Char
                Dim initialsStr As String = ""

                ThisForm.Caption.ToCharArray.CopyTo(captionChars, 0)
                initialsStr += captionChars(0).ToString()

                For i As Integer = 1 To (captionChars.Length - 1)
                    If captionChars(i) = " " Then
                        initialsStr += captionChars(i + 1).ToString()
                    End If
                Next

                ThisForm.Variables("varFormInitials").SetValue(initialsStr)
            End If
        End Sub

        Sub SetDynamicFormCaption()
            If ThisForm.Variables("parm_DisplayTransactionNumInTitleBar").GetValueOfByte(0) = 1 Then
                Dim transactionNum As String
                transactionNum = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")
                Dim formInitials As String = ThisForm.Variables("varFormInitials").GetValueOfString("")
                ThisForm.Caption = "~LIT~(" + formInitials + " - " + transactionNum + ")"
            End If
        End Sub

        Sub ResetDynamicCaption()
            ThisForm.Caption = "fCustomerOrderLines"
        End Sub

        Function ValidateRefType() As Integer
            ValidateRefType = 0
            If ThisForm.Components("RefTypeEdit").ValidateData(True) = False Then
                ValidateRefType = -1
            End If
        End Function

        Sub ValidatePLOverrideItmDescInInvoiceChk()
            Dim IsCSIB_79704Active As String = Application.Variables("IsCSIB_79704ActiveVar").Value
            Dim IsPLActive As String = Application.Variables("Avail_PL").Value
            Dim IsPLOverrideItmDescInInvoiceChecked As String = ThisForm.PrimaryIDOCollection().GetCurrentObjectProperty("PLOverrideItemDescription")

            If IsCSIB_79704Active = "1" And IsPLActive = "1" Then
                If IsPLOverrideItmDescInInvoiceChecked = "1" Then
                    ThisForm.Components("PLAlternateItemDescEdit").Enabled = True
                    ThisForm.Components("PLAlternateItemDescriptionGridCol").Enabled = True
                Else
                    ThisForm.Components("PLAlternateItemDescEdit").Enabled = False
                    ThisForm.Components("PLAlternateItemDescEdit").Value = ""
                    ThisForm.Components("PLAlternateItemDescriptionGridCol").Enabled = False
                    ThisForm.Components("PLAlternateItemDescriptionGridCol").Value = ""
                End If
            End If
        End Sub

        Function SetDueDate() As Integer
            If IsNumeric(ThisForm.Variables("ICDuePeriod").Value) Then
                ThisForm.Components("DueDateEdit").SetValue(DateAdd("d", CLng(ThisForm.Variables("ICDuePeriod").Value), ThisForm.Components("CoOrderDateEdit").GetValueOfDateTime(Date.MinValue)))
                ThisForm.Components("DueDateGridCol").SetValue(DateAdd("d", CLng(ThisForm.Variables("ICDuePeriod").Value), ThisForm.Components("CoOrderDateGridCol").GetValueOfDateTime(Date.MinValue)))
            End If
        End Function

        Sub ReceiveIBCMessage()
            Dim messageContext As String 
            Dim messageData As String 
            Dim strItem As String
            Dim apiResponse As InvokeResponseData

            messageContext= ThisForm.Variables("StdWebPageMessageContext").Value
            If messageContext <> "applicationSetFields"  Then Return
            messageData  = ThisForm.Variables("StdWebPageMessageData").Value
            If String.IsNullOrEmpty(messageData) Then Return

            'CALL API to parse messageContext to get item
            apiResponse = IDOClient.Invoke("SLFormExtMsgEntities", "GetRecommendedItem", messageData, strItem)
            If Not apiResponse.Parameters(1).IsNull Then strItem =apiResponse.Parameters(1).GetValue(Of String)()
            If String.IsNullOrEmpty(strItem) Then Return

            ThisForm.PrimaryIDOCollection.New()
            ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("Item", strItem)
            ThisForm.Components("ItemEdit").ValidateData(True)
            ThisForm.GenerateEvent("ItemChanged")
            ThisForm.Variables("StdWebPageMessageContext").Value = ""
            ThisForm.Variables("StdWebPageMessageData").Value = ""
        End Sub

        Sub ValidateCPQWhseChange()
            Dim bRS7786_5 As Boolean     
            Dim sParam As String
            bRS7786_5 = Application.Variables("IsRS7786_5ActiveVar").GetValueOfInteger(0) = 1 
            sParam = Application.GetStringValue("@sValidateWhseCPQMP")
            If bRS7786_5 and Application.Variables("Avail_Cfg").Value = "1" And ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("ItOrderConfigurable") = "1" Then
               If ThisForm.Components("DerLinePlantEdit").value <> ThisForm.Components("PrevPlantEdit").value Then 
                  If ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerCfgMPIsPending") = "0" Then
                      ThisForm.CallGlobalScript("MsgApp", "Clear", "NoPrompt", "SuccessFailure",
                        "mI=Changed0", "@sPlant", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                     ThisForm.CallGlobalScript("MsgApp", "NoClear", "Prompt", "SuccessFailure",
                        "mW=ReconfigureItem", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                  End If            
               End If              
            End If
        End Sub

        Sub SetDueDateProperty()
            If Application.Variables("Avail_AU").Value <> "1" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("DueDate",
                      ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerDueDate"))
            End If
        End Sub

        Sub SetPromiseDateProperty()
            If Application.Variables("Avail_AU").Value <> "1" Then
                ThisForm.PrimaryIDOCollection.SetCurrentObjectPropertyPlusModifyRefresh("PromiseDate",
                      ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("DerPromiseDate"))
            End If
        End Sub

        Sub ApplyAllowNonItemVar()
            Dim bFranceCP As Boolean
            Dim bCSIB_97770 As Boolean
            Dim parmAllowNonItem As String

            bCSIB_97770 = Application.Variables("IsCSIB_97770ActiveVar").GetValueOfInteger(0) = 1
            bFranceCP = Application.Variables("Avail_FR").GetValueOfInteger(0) = 1
            parmAllowNonItem = Application.Variables("Parm_FRAllowNonItemsOnOrders").Value

            If bCSIB_97770 And bFranceCP Then
                If Not String.IsNullOrEmpty(parmAllowNonItem) Then
                    ThisForm.Variables("AllowNonItemVar").SetValue(parmAllowNonItem)
                Else
                    ThisForm.GenerateEvent("FRSetAllowNonItemVar")
                End If
            End If
        End Sub

        Sub SetFRNonInventoryDueDate()
            If Not String.IsNullOrEmpty(ThisForm.Variables("FRNonInvItemDuePeriod").Value) AndAlso CLng(ThisForm.Variables("FRNonInvItemDuePeriod").Value) <> 0 Then
               ThisForm.Components("DueDateEdit").SetValue(DateAdd("d", CLng(ThisForm.Variables("FRNonInvItemDuePeriod").Value), ThisForm.Components("CoOrderDateEdit").GetValueOfDateTime(Date.MinValue)))
               ThisForm.Components("DueDateGridCol").SetValue(DateAdd("d", CLng(ThisForm.Variables("FRNonInvItemDuePeriod").Value), ThisForm.Components("CoOrderDateGridCol").GetValueOfDateTime(Date.MinValue)))
            End If
        End Sub

        ' Custom button click to open external N8N form with Customer Order number
        Sub ConfigureN8nButtonClick()
            Dim coNum As String = ThisForm.PrimaryIDOCollection.GetCurrentObjectProperty("CoNum")
            
            If Not String.IsNullOrEmpty(coNum) Then
                Dim encodedCoNum As String = System.Uri.EscapeDataString(coNum)
                Dim fullUrl As String = "https://n8n.bainultra.dev/form/a35b0cad-1d7c-439f-a1e5-744311054b2a?orderNumber=" & encodedCoNum
                ThisForm.Variables("N8nUrl").SetValue(fullUrl)
                ThisForm.GenerateEvent("OpenN8nUrl")
            Else
                ThisForm.CallGlobalScript("MsgApp", "Clear", "Prompt", "SuccessFailure", _
                    "mE=CmdInvalid", "@sSelectCustomerOrderFirst", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End Sub

    End Class
End Namespace
