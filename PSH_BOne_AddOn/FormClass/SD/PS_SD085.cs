using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 입금조회
    /// </summary>
    internal class PS_SD085 : PSH_BaseClass
    {
        public string oFormUniqueID;
        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_USERDS01; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        //private int oLast_Mode;
        //private int oSeq;

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD085.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD085_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD085");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);

                PS_SD085_CreateItems();
                PS_SD085_SetComboBox();
                //PS_SD085_Initialize();

                oForm.EnableMenu("1281", false); //찾기
                oForm.EnableMenu("1282", false); //추가
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD085_CreateItems()
        {
            try
            {
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.AddMonths(-1).ToString("yyyyMM01");

                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_SD085_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    int ErrNum = 0;
        //    object TempForm01 = null;
        //    SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //    string ItemType = null;
        //    string RequestDate = null;
        //    string Size = null;
        //    string ItemCode = null;
        //    string ItemName = null;
        //    string Unit = null;
        //    string DueDate = null;
        //    string RequestNo = null;
        //    int Qty = 0;
        //    decimal Weight = default(decimal);
        //    string RFC_Sender = null;
        //    double Calculate_Weight = 0;
        //    int Seq = 0;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                break;
        //            //et_KEY_DOWN ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                if (pVal.CharPressed == 9)
        //                {
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "Btn01")
        //                {
        //                    PS_SD085_LoadData();
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                break;
        //            //et_COMBO_SELECT ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                ////5
        //                if (pVal.ItemChanged == true)
        //                {
        //                    if (pVal.ItemUID == "BPLId")
        //                    {
        //                        PS_SD085_FlushToItemValue(pVal.ItemUID);
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CLICK:
        //                ////6
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //                ////7
        //                break;
        //            //et_MATRIX_LINK_PRESSED /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                ////8
        //                break;
        //            //                If pVal.ItemUID = "Mat01" And pVal.ColUID = "TrandId" Then
        //            //                   'Set TempForm01 = New "392"
        //            //                ElseIf pVal.ItemUID = "Mat01" And pVal.ColUID = "Ref1" Then
        //            //                        Set TempForm01 = New PS_PP040
        //            //                End If
        //            //
        //            //                Call TempForm01.LoadForm(oMat01.Columns("DocEntry").Cells(pVal.Row).Specific.Value)
        //            //                Set TempForm01 = Nothing

        //            //et_VALIDATE ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                if (pVal.ItemChanged == true)
        //                {
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        PS_SD085_FlushToItemValue(pVal.ItemUID);
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //                ////11
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //                ////18
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //                ////19
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //                ////20
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                ////27
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //                ////3
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //                ////4
        //                break;
        //            //et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //                ////17
        //                SubMain.RemoveForms(oFormUniqueID);
        //                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oForm = null;
        //                //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oMat01 = null;
        //                //UPGRADE_NOTE: oDS_PS_USERDS01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PS_USERDS01 = null;
        //                break;
        //        }
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgressBar01 = null;
        //    if (ErrNum == 101)
        //    {
        //        ErrNum = 0;
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //        BubbleEvent = false;
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pVal.BeforeAction == false))
        //    {
        //        switch (pVal.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                break;
        //        }
        //    }
        //    return;
        //Raise_MenuEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((BusinessObjectInfo.BeforeAction == true))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((BusinessObjectInfo.BeforeAction == false))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //    }
        //    return;
        //Raise_FormDataEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if ((eventInfo.BeforeAction == true))
        //    {
        //        ////작업
        //    }
        //    else if ((eventInfo.BeforeAction == false))
        //    {
        //        ////작업
        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion



        #region PS_SD085_Initialize
        //public void PS_SD085_Initialize()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbouiCOM.ComboBox oCombo = null;

        //    ////아이디별 사업장 세팅
        //    oCombo = oForm.Items.Item("BPLId").Specific;
        //    oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);


        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    return;
        //PS_SD085_Initialize_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD085_Initialize_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD085_FlushToItemValue
        //private void PS_SD085_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short ErrNum = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    string CardCode = null;

        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    switch (oUID)
        //    {
        //        case "CardCode":
        //            oForm.Freeze(true);
        //            if (oUID == "CarCode")
        //            {
        //                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                sQry = "Select CardName From OCRD Where CardCode = '" + Strings.Trim(oForm.Items.Item("CardCode").Specific.Value) + "'";
        //                oRecordSet01.DoQuery(sQry);

        //                //UPGRADE_WARNING: oForm.Items(CardName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm.Items.Item("CardName").Specific.Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //            }
        //            oForm.Freeze(false);
        //            break;
        //    }

        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    return;
        //PS_SD085_FlushToItemValue_Error:
        //    oForm.Freeze(false);
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD085_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion


        #region PS_SD085_LoadData
        //public void PS_SD085_LoadData()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    short ErrNum = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    string DocDateTo = null;
        //    string BPLId = null;
        //    string DocDateFr = null;
        //    string CardCode = null;

        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    oMat01.Clear();
        //    oDS_PS_USERDS01.Clear();

        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocDateFr = Strings.Trim(oForm.Items.Item("DocDateFr").Specific.Value);
        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocDateTo = Strings.Trim(oForm.Items.Item("DocDateTo").Specific.Value);
        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);

        //    if (string.IsNullOrEmpty(BPLId))
        //        BPLId = "%";
        //    if (string.IsNullOrEmpty(DocDateFr))
        //        DocDateFr = "18990101";
        //    if (string.IsNullOrEmpty(DocDateTo))
        //        DocDateTo = "20991231";
        //    if (string.IsNullOrEmpty(CardCode))
        //        CardCode = "%";

        //    sQry = "EXEC [PS_SD085_01] '" + BPLId + "', '" + DocDateFr + "', '" + DocDateTo + "', '" + CardCode + "'";
        //    oRecordSet01.DoQuery(sQry);

        //    if (oRecordSet01.RecordCount == 0)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
        //        //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        oRecordSet01 = null;
        //        oForm.Freeze(false);
        //        return;
        //    }

        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //    for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //    {
        //        if (i + 1 > oDS_PS_USERDS01.Size)
        //        {
        //            oDS_PS_USERDS01.InsertRecord((i));
        //        }

        //        oMat01.AddRow();
        //        oDS_PS_USERDS01.Offset = i;
        //        oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //        oDS_PS_USERDS01.SetValue("U_ColDt01", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oRecordSet01.Fields.Item("DocDate").Value), "YYYYMMDD"));
        //        oDS_PS_USERDS01.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("TransId").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("Ref1").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("CardCode").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet01.Fields.Item("CardName").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet01.Fields.Item("LineMemo").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg06", i, Strings.Trim(oRecordSet01.Fields.Item("Account").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("Amt").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("RefAmt").Value));
        //        oDS_PS_USERDS01.SetValue("U_ColReg07", i, Strings.Trim(oRecordSet01.Fields.Item("RefNum").Value));
        //        //----------------------------------------------------------------------------------------------------------
        //        oRecordSet01.MoveNext();
        //        ProgBar01.Value = ProgBar01.Value + 1;
        //        ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
        //    }
        //    oMat01.LoadFromDataSource();
        //    //            oMat01.AutoResizeColumns
        //    ProgBar01.Stop();
        //    oForm.Freeze(false);


        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    return;
        //PS_SD085_LoadData_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    oForm.Freeze(false);
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD085_LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion
    }
}
