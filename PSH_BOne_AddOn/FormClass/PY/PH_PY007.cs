using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
//using Microsoft.Office.Interop;
using SAPbobsCOM;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    internal class PH_PY007 : PSH_BaseClass
    {
        ////********************************************************************************
        ////  File           : PH_PY007.cls
        ////  Module         : 기타관리 > 사용/공용/출장
        ////  Desc           : 유류단가등록
        ////********************************************************************************

        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY007A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY007B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        public override void LoadForm(string oFormDocEntry01 = "")
        {

            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY007.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue.ToString() + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY007_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY007");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(ref strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                ////fm_ADD_MODE
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY007_CreateItems();
                PH_PY007_EnableMenus();
                PH_PY007_SetDocument(oFormDocEntry01);
                //    Call PH_PY007_FormResize
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        private void PH_PY007_CreateItems()
        {
            string sQry = string.Empty;

            SAPbouiCOM.ComboBox oCombo = null;
         
            SAPbobsCOM.Recordset oRecordSet = null;

            oForm.Freeze(true);

            try
            {
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oDS_PH_PY007A = oForm.DataSources.DBDataSources.Item("@PH_PY007A");
                oDS_PH_PY007B = oForm.DataSources.DBDataSources.Item("@PH_PY007B");
                
                oMat1 = oForm.Items.Item("Mat01").Specific;
                ////@PH_PY007B

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();


                ////----------------------------------------------------------------------------------------------
                //// 기본사항
                ////----------------------------------------------------------------------------------------------

                ////사업장
                oCombo = oForm.Items.Item("CLTCOD").Specific;
                //    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
                //    Call SetReDataCombo(oForm, sQry, oCombo)
                //
                //    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
                //    oCombo.Select CLTCOD, psk_ByValue

                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //// 년
                oDS_PH_PY007A.SetValue("U_Year", 0, DateTime.Now.ToString("yyyy"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY007_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);

                if (oForm.Visible == false)
                {
                    oForm.Visible = true;
                }

                oForm.Update();
                //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        private void PH_PY007_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY007_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void PH_PY007_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry01)))
                {
                    PH_PY007_FormItemEnabled();
                    //        Call PH_PY007_AddMatrixRow
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY007_FormItemEnabled();
                    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    oForm.Items.Item("Code").Specific.VALUE = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            
        }

        public void PH_PY007_FormItemEnabled()
        {
            oForm.Freeze(true);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", false);
                    ////문서추가
                    oForm.EnableMenu("1293", false);
                    ////행삭제

                    oForm.Items.Item("Year").Enabled = true;

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가
                    oForm.EnableMenu("1293", false);
                    ////행삭제

                    oForm.Items.Item("Year").Enabled = true;
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가
                    oForm.EnableMenu("1293", false);
                    ////행삭제

                    oForm.Items.Item("Year").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY007_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        ////1

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    if (PH_PY007_DataValidCheck("Y") == false)
                                    {
                                        BubbleEvent = false;
                                    }
                                    ////해야할일 작업
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    ////해야할일 작업

                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {

                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    PH_PY007_FormItemEnabled();
                                }
                            }

                        }
                        else if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    if (pVal.ActionSuccess == true)
                                    {
                                        PH_PY007_FormItemEnabled();
                                        //                        Call PH_PY007_AddMatrixRow
                                    }
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    if (pVal.ActionSuccess == true)
                                    {
                                        PH_PY007_FormItemEnabled();
                                        //                        Call PH_PY007_AddMatrixRow
                                    }
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    if (pVal.ActionSuccess == true)
                                    {
                                        PH_PY007_FormItemEnabled();
                                    }
                                }
                            }

                        }
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        ////2
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        ////3
                        switch (pVal.ItemUID)
                        {
                            case "Mat01":
                                if (pVal.Row > 0)
                                {
                                    oLastItemUID = pVal.ItemUID;
                                    oLastColUID = pVal.ColUID;
                                    oLastColRow = pVal.Row;
                                }
                                break;
                            default:
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = "";
                                oLastColRow = 0;
                                break;
                        }
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        ////4
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        ////5
                        oForm.Freeze(true);
                        if (pVal.BeforeAction == true)
                        {

                        }
                        else if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemChanged == true)
                            {

                            }
                        }
                        oForm.Freeze(false);
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        ////6
                        if (pVal.BeforeAction == true)
                        {
                            switch (pVal.ItemUID)
                            {
                                case "Mat01":
                                    if (pVal.Row > 0)
                                    {
                                        oMat1.SelectRow(pVal.Row, true, false);
                                    }
                                    break;
                            }

                            switch (pVal.ItemUID)
                            {
                                case "Mat01":
                                    if (pVal.Row > 0)
                                    {
                                        oLastItemUID = pVal.ItemUID;
                                        oLastColUID = pVal.ColUID;
                                        oLastColRow = pVal.Row;
                                    }
                                    break;
                                default:
                                    oLastItemUID = pVal.ItemUID;
                                    oLastColUID = "";
                                    oLastColRow = 0;
                                    break;
                            }
                        }
                        else if (pVal.BeforeAction == false)
                        {

                        }
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        ////7
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        ////8
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                        ////9
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        ////10
                        oForm.Freeze(true);
                        if (pVal.BeforeAction == true)
                        {

                        }
                        else if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemChanged == true)
                            {
                                if (pVal.ItemUID == "Year")
                                {
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        PH_PY007_Create_MonthData();
                                    }
                                }
                            }
                        }
                        oForm.Freeze(false);
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
                        ////11
                        if (pVal.BeforeAction == true)
                        {
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            oMat1.LoadFromDataSource();
                            PH_PY007_FormItemEnabled();
                        }
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                        ////12
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        ////16
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
                        ////17
                        if (pVal.BeforeAction == true)
                        {
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            SubMain.Remove_Forms(oFormUniqueID);
                            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            oForm = null;
                            //UPGRADE_NOTE: oDS_PH_PY007A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            oDS_PH_PY007A = null;
                            //UPGRADE_NOTE: oDS_PH_PY007B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            oDS_PH_PY007B = null;

                            //UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            oMat1 = null;

                        }
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        ////18
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                        ////19
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        ////20
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        ////21
                        if (pVal.BeforeAction == true)
                        {

                        }
                        else if (pVal.BeforeAction == false)
                        {
                            //                oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10
                            //                oForm.Items.Item("79").Height = oForm.Items.Item("80").Height
                            //
                            //                oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16
                            //                oForm.Items.Item("77").Height = oForm.Items.Item("78").Height
                            //
                            //                oForm.Items.Item("8").Width = oForm.Items.Item("Mat02").Left + oForm.Items.Item("Mat02").Width + 5
                            //                oForm.Items.Item("8").Height = oForm.Items.Item("1").Top - 80
                        }
                        break;
                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                        ////22
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                        ////23
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        ////27
                        if (pVal.BeforeAction == true)
                        {

                        }
                        else if (pVal.Before_Action == false)
                        {
                            //                If pVal.ItemUID = "Code" Then
                            //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY007A", "Code")
                            //                End If
                        }
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                        ////37
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                        ////38
                        break;

                    //----------------------------------------------------------
                    case SAPbouiCOM.BoEventTypes.et_Drag:
                        ////39
                        break;

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_Raise_ItemEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            DataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY007A", "Code");
                            ////접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY007_FormItemEnabled();
                            break;
                        //                Call PH_PY007_AddMatrixRow
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY007_FormItemEnabled();
                            //                Call PH_PY007_AddMatrixRow
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY007_FormItemEnabled();
                            break;
                        //                Call PH_PY007_AddMatrixRow
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY007_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
                            break;

                            //                Call PH_PY007_AddMatrixRow
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public void PH_PY007_AddMatrixRow()
        {
            int oRow = 0;
            
            oForm.Freeze(true);

            try
            {
                ////[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY007B.GetValue("U_Date", oRow - 1))))
                    {
                        if (oDS_PH_PY007B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY007B.InsertRecord((oRow));
                        }
                        oDS_PH_PY007B.Offset = oRow;
                        oDS_PH_PY007B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY007B.SetValue("U_Month", oRow, "");
                        oDS_PH_PY007B.SetValue("U_Gasoline", oRow, "");
                        oDS_PH_PY007B.SetValue("U_Diesel", oRow, Convert.ToString(0));
                        oDS_PH_PY007B.SetValue("U_LPG", oRow, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY007B.Offset = oRow - 1;
                        oDS_PH_PY007B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY007B.SetValue("U_Month", oRow - 1, "");
                        oDS_PH_PY007B.SetValue("U_Gasoline", oRow - 1, "");
                        oDS_PH_PY007B.SetValue("U_Diesel", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY007B.SetValue("U_LPG", oRow - 1, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY007B.Offset = oRow;
                    oDS_PH_PY007B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY007B.SetValue("U_Month", oRow, "");
                    oDS_PH_PY007B.SetValue("U_Gasoline", oRow, "");
                    oDS_PH_PY007B.SetValue("U_Diesel", oRow, Convert.ToString(0));
                    oDS_PH_PY007B.SetValue("U_LPG", oRow, Convert.ToString(0));
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public void PH_PY007_FormClear()
        {
            string DocEntry = string.Empty;

            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY007'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {                    
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public bool PH_PY007_DataValidCheck(string ChkYN)
        {
            bool functionReturnValue = false;         
            functionReturnValue = false;
                        
            string sQry = string.Empty;
            string tCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = null;

            try
            {
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (ChkYN == "Y")
                {
                    if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY007A.GetValue("U_CLTCOD", 0))))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }

                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY007A.GetValue("U_Year", 0))))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }


                //// 코드,이름 저장
                tCode = Strings.Trim(oDS_PH_PY007A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY007A.GetValue("U_Year", 0));
                oDS_PH_PY007A.SetValue("Code", 0, tCode);
                oDS_PH_PY007A.SetValue("Name", 0, tCode);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //// 데이터 중복 체크
                    sQry = "SELECT Code FROM [@PH_PY007A] WHERE Code = '" + tCode + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 데이터가 존재합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }

                if (ChkYN == "Y")
                {
                    if (oMat1.VisualRowCount == 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 없습니다. 월을생성을 하기바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY007_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        private void PH_PY007_Create_MonthData()
        {            
            int i = 0;         

            oMat1.LoadFromDataSource();

            try
            {
                for (i = 0; i <= 11; i++)
                {
                    if (i + 1 > oDS_PH_PY007B.Size)
                    {
                        oDS_PH_PY007B.InsertRecord((i));
                    }
                    oDS_PH_PY007B.Offset = i;
                    oDS_PH_PY007B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY007B.SetValue("U_Month", i, Strings.Right("0" + Convert.ToString(i + 1), 2));
                    oDS_PH_PY007B.SetValue("U_Gasoline", i, Convert.ToString(0));
                    oDS_PH_PY007B.SetValue("U_Diesel", i, Convert.ToString(0));
                    oDS_PH_PY007B.SetValue("U_LPG", i, Convert.ToString(0));
                }

                oMat1.LoadFromDataSource();
                //    Call PH_PY007_AddMatrixRow
                //    oMat1.Columns("Col1").Cells(i + 1).CLICK ct_Regular

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_Create_MonthData_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public bool PH_PY007_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (DataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY007A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    goto PH_PY007_Validate_Exit;
                }
                //
                if (ValidateType == "수정")
                {

                }
                else if (ValidateType == "행삭제")
                {

                }
                else if (ValidateType == "취소")
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY007_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);            
                functionReturnValue = false;
            }
            
        PH_PY007_Validate_Exit:
            return functionReturnValue;
        }
    }
}