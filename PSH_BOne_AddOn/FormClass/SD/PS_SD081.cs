using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 거래처별 한도 조회
    /// </summary>
    internal class PS_SD081 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_SD081H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_SD081L; //등록라인
        private SAPbouiCOM.Form oBaseForm; ////부모폼
        //private string oBaseItemUID01;
        //private string oBaseColUID01;
        private int oBaseColRow;
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pBaseForm"></param>
        /// <param name="pBaseColRow"></param>
        public void LoadForm(SAPbouiCOM.Form pBaseForm, int pBaseColRow)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD081.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SD081_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SD081");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                oBaseForm = pBaseForm;
                oBaseColRow = pBaseColRow;

                //PS_SD081_CreateItems();
                //PS_SD081_SetComboBox();
                //PS_SD081_Initialize();
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


        //private void PS_SD081_CreateItems()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    string oQuery01 = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    oDS_PS_SD081L = oForm.DataSources.DBDataSources("@PS_USERDS01");
        //    oMat01 = oForm.Items.Item("Mat01").Specific;
        //    oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
        //    oMat01.AutoResizeColumns();

        //    oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
        //    //    oForm.DataSources.UserDataSources.Item("BPLId").Value = oBaseForm.Items("BPLId").Specific.Value

        //    oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
        //    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");
        //    //    oForm.DataSources.UserDataSources.Item("CardCode").Value = oBaseForm.Items("CardCode").Specific.Value

        //    oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
        //    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");
        //    //    oForm.DataSources.UserDataSources.Item("CardName").Value = oBaseForm.Items("CardCode").Specific.Value

        //    oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
        //    //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
        //    oForm.DataSources.UserDataSources.Item("DocDate").Value = Convert.ToString(DateAndTime.Today);
        //    //Format(Now, "YYYY-MM") & "-01"


        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD081_CreateItems_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "PS_SD081_CreateItems_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}


        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    int ErrNum = 0;
        //    object TempForm01 = null;

        //    ////BeforeAction = True
        //    if ((pVal.BeforeAction == true))
        //    {
        //        switch (pVal.EventType)
        //        {
        //            //et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                ////1
        //                if (pVal.ItemUID == "Btn01")
        //                {
        //                    PS_SD081_SetBaseForm();
        //                }
        //                else if (pVal.ItemUID == "Btn02")
        //                {
        //                    PS_SD081_LoadData();
        //                }
        //                break;
        //            //et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
        //                if (pVal.CharPressed == 9)
        //                {
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
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
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //                ////2
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
        //            //et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //            case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //                ////10
        //                if (pVal.ItemChanged == true)
        //                {
        //                    if (pVal.ItemUID == "CardCode")
        //                    {
        //                        PS_SD081_FlushToItemValue(pVal.ItemUID);
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
        //                //UPGRADE_NOTE: oDS_PS_SD081H 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PS_SD081H = null;
        //                //UPGRADE_NOTE: oDS_PS_SD081L 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PS_SD081L = null;
        //                break;
        //        }
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

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
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pVal.BeforeAction == true)
        //    {

        //    }
        //    else if (pVal.BeforeAction == false)
        //    {

        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion



        #region PS_SD081_SetComboBox
        //public void PS_SD081_SetComboBox()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////콤보에 기본값설정
        //    SAPbouiCOM.ComboBox oCombo = null;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;

        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    //// 사업장
        //    oCombo = oForm.Items.Item("BPLId").Specific;
        //    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
        //    oRecordSet01.DoQuery(sQry);
        //    while (!(oRecordSet01.EoF))
        //    {
        //        oCombo.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
        //        oRecordSet01.MoveNext();
        //    }
        //    oCombo.Select("4", SAPbouiCOM.BoSearchKey.psk_ByValue);

        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD081_SetComboBox_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD081_SetComboBox_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD081_Initialize
        //public void PS_SD081_Initialize()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbouiCOM.ComboBox oCombo = null;

        //    ////아이디별 사업장 세팅
        //    oCombo = oForm.Items.Item("BPLId").Specific;
        //    oCombo.Select(MDC_PS_Common.User_BPLId(), SAPbouiCOM.BoSearchKey.psk_ByValue);

        //    ////아이디별 사번 세팅
        //    //    oForm.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD

        //    ////아이디별 부서 세팅
        //    //    Set oCombo = oForm.Items("DeptCode").Specific
        //    //    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    return;
        //PS_SD081_Initialize_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD081_Initialize_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD081_FlushToItemValue
        //private void PS_SD081_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    short ErrNum = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;

        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    switch (oUID)
        //    {
        //        case "CardCode":
        //            sQry = "Select CardName From [OCRD] Where CardCode = '" + Strings.Trim(oForm.DataSources.UserDataSources.Item("CardCode").Value) + "'";
        //            oRecordSet01.DoQuery(sQry);

        //            oForm.DataSources.UserDataSources.Item("CardName").Value = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //            break;
        //    }

        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD081_FlushToItemValue_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD081_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD081_LoadData
        //public void PS_SD081_LoadData()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    short i = 0;
        //    string sQry = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string BPLId = null;
        //    string CardCode = null;
        //    object DocDate = null;

        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    BPLId = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    //UPGRADE_WARNING: DocDate 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);


        //    if (string.IsNullOrEmpty(BPLId))
        //        BPLId = "%";
        //    if (string.IsNullOrEmpty(CardCode))
        //        CardCode = "%";

        //    //UPGRADE_WARNING: DocDate 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    if (string.IsNullOrEmpty(DocDate))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "기준일자가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
        //        return;
        //    }


        //    //UPGRADE_WARNING: DocDate 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    sQry = "EXEC [S139_hando] '" + CardCode + "','" + DocDate + "'";
        //    oRecordSet01.DoQuery(sQry);

        //    oMat01.Clear();
        //    oDS_PS_SD081L.Clear();

        //    if (oRecordSet01.RecordCount == 0)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.:" + Err().Number + " - " + Err().Description, ref "W");
        //        //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        oRecordSet01 = null;
        //        return;
        //    }

        //    oForm.Freeze(true);
        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

        //    for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //    {
        //        if (i + 1 > oDS_PS_SD081L.Size)
        //        {
        //            oDS_PS_SD081L.InsertRecord((i));
        //        }

        //        oMat01.AddRow();
        //        oDS_PS_SD081L.Offset = i;
        //        oDS_PS_SD081L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //        oDS_PS_SD081L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("CardCode").Value));
        //        oDS_PS_SD081L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("CardName").Value));
        //        oDS_PS_SD081L.SetValue("U_ColSum01", i, Strings.Trim(oRecordSet01.Fields.Item("U_CreditP").Value));
        //        //현재여신 금액
        //        oDS_PS_SD081L.SetValue("U_ColSum02", i, Strings.Trim(oRecordSet01.Fields.Item("U_MiSuP").Value));
        //        //미수계
        //        oDS_PS_SD081L.SetValue("U_ColSum06", i, Strings.Trim(oRecordSet01.Fields.Item("U_ArfAmt").Value));
        //        //어음
        //        oDS_PS_SD081L.SetValue("U_ColSum07", i, Strings.Trim(oRecordSet01.Fields.Item("U_Budo").Value));
        //        //부도
        //        oDS_PS_SD081L.SetValue("U_ColSum08", i, Strings.Trim(oRecordSet01.Fields.Item("U_MisuTot").Value));
        //        //채권계

        //        //Trim (oRecordSet01.Fields("U_MiSuP").Value) '미수금액
        //        //Trim(oRecordSet01.Fields("U_ArfAmt").Value) '어음잔액
        //        //Trim(oRecordSet01.Fields("U_Budo").Value) '부도어음

        //        oDS_PS_SD081L.SetValue("U_ColSum03", i, Strings.Trim(oRecordSet01.Fields.Item("U_Balance").Value));
        //        oDS_PS_SD081L.SetValue("U_ColSum04", i, Strings.Trim(oRecordSet01.Fields.Item("U_OutPreP").Value));
        //        oDS_PS_SD081L.SetValue("U_ColSum05", i, Strings.Trim(oRecordSet01.Fields.Item("OverAmt").Value));

        //        oRecordSet01.MoveNext();
        //        ProgBar01.Value = ProgBar01.Value + 1;
        //        ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
        //    }
        //    oMat01.LoadFromDataSource();
        //    oMat01.AutoResizeColumns();
        //    ProgBar01.Stop();
        //    oForm.Freeze(false);

        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    return;
        //PS_SD081_LoadData_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet01 = null;
        //    MDC_Com.MDC_GF_Message(ref "PS_SD081_LoadData_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion

        #region PS_SD081_SetBaseForm
        //private void PS_SD081_SetBaseForm()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    int j = 0;
        //    int ErrNum = 0;
        //    int sRow = 0;

        //    SAPbouiCOM.Matrix oBaseMat01 = null;
        //    SAPbouiCOM.DBDataSource oBaseDS_PS_SD080L = null;
        //    oBaseMat01 = oBaseForm.Items.Item("Mat01").Specific;
        //    oBaseDS_PS_SD080L = oBaseForm.DataSources.DBDataSources("@PS_SD080L");

        //    oBaseForm.Freeze(true);
        //    oBaseMat01.Clear();
        //    oBaseMat01.FlushToDataSource();
        //    oBaseMat01.LoadFromDataSource();

        //    oMat01.FlushToDataSource();
        //    sRow = 0;
        //    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //    {
        //        if (Strings.Trim(oDS_PS_SD081L.GetValue("U_ColReg01", i)) == "Y")
        //        {
        //            if (sRow + 1 > oBaseDS_PS_SD080L.Size)
        //            {
        //                oBaseDS_PS_SD080L.InsertRecord((sRow));
        //            }

        //            oBaseMat01.AddRow();
        //            oBaseDS_PS_SD080L.Offset = sRow;
        //            oBaseDS_PS_SD080L.SetValue("U_LineNum", sRow, Convert.ToString(sRow + 1));
        //            oBaseDS_PS_SD080L.SetValue("U_CardCode", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColReg02", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_CardName", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColReg03", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_CreditP", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum01", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_MiSuP", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum02", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_Balance", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum03", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_Bill", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum06", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_Budo", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum07", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_TotAmt", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum08", i)));

        //            oBaseDS_PS_SD080L.SetValue("U_OutPreP", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum04", i)));
        //            oBaseDS_PS_SD080L.SetValue("U_RequestP", sRow, Strings.Trim(oDS_PS_SD081L.GetValue("U_ColSum05", i)));


        //            oBaseColRow = oBaseColRow + 1;
        //            sRow = sRow + 1;
        //        }
        //    }

        //    oBaseMat01.LoadFromDataSource();
        //    oBaseForm.Freeze(false);
        //    oForm.Close();
        //    return;
        //PS_SD081_SetBaseForm_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref " ", ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "PS_SD081_SetBaseForm_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //}
        #endregion
    }
}
