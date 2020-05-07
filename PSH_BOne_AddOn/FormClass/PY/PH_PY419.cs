using Microsoft.VisualBasic;
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사업장정보등록
    /// </summary>
    internal class PH_PY419 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY419;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 사업장정보등록
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY419.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY419_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY419");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //    oForm.DataBrowser.BrowseBy = "Code"

                oForm.Freeze(true);
                PH_PY419_CreateItems();
                PH_PY419_EnableMenus();
                PH_PY419_SetDocument(oFromDocEntry01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY419_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY419");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY419");
                oDS_PH_PY419 = oForm.DataSources.DataTables.Item("PH_PY419");

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                ////년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                ////사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
                ////성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY419_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        //		private void PH_PY419_EnableMenus()
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.EnableMenu("1283", false);
        //			////제거
        //			oForm.EnableMenu("1284", false);
        //			////취소
        //			oForm.EnableMenu("1293", false);
        //			////행삭제

        //			return;
        //			PH_PY419_EnableMenus_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY419_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY419_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY419_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY202_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        //private void PH_PY419_SetDocument(string oFromDocEntry01)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement


        //    if ((string.IsNullOrEmpty(oFromDocEntry01)))
        //    {
        //        PH_PY419_FormItemEnabled();
        //        //        Call PH_PY419_AddMatrixRow
        //    }
        //    else
        //    {
        //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //        PH_PY419_FormItemEnabled();
        //        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
        //        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //    }
        //    return;
        //PH_PY419_SetDocument_Error:

        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY419_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        //public void PH_PY419_FormItemEnabled()
        //{
        //    SAPbouiCOM.ComboBox oCombo = null;
        //    string sQry = null;
        //    int i = 0;
        //    SAPbobsCOM.Recordset oRecordSet = null;


        //    string CLTCOD = null;
        //    string sPosDate = null;

        //    // ERROR: Not supported in C#: OnErrorStatement

        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    oForm.Freeze(true);
        //    if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //    {


        //        oForm.EnableMenu("1281", false);
        //        ////문서찾기
        //        oForm.EnableMenu("1282", true);
        //        ////문서추가

        //        //UPGRADE_WARNING: oForm.Items(Year).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("Year").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;
        //        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("MSTCOD").Specific.VALUE = "";
        //        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oForm.Items.Item("FullName").Specific.VALUE = "";

        //        //// 접속자에 따른 권한별 사업장 콤보박스세팅
        //        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");


        //    }
        //    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
        //    {
        //        //// 접속자에 따른 권한별 사업장 콤보박스세팅
        //        MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

        //        oForm.EnableMenu("1281", false);
        //        ////문서찾기
        //        oForm.EnableMenu("1282", true);
        //        ////문서추가


        //    }
        //    else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
        //    {
        //        //// 접속자에 따른 권한별 사업장 콤보박스세팅
        //        MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

        //        oForm.EnableMenu("1281", true);
        //        ////문서찾기
        //        oForm.EnableMenu("1282", true);
        //        ////문서추가

        //    }
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    oForm.Freeze(false);
        //    return;
        //PH_PY419_FormItemEnabled_Error:

        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    oForm.Freeze(false);
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}


        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID01);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY419);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        //if (PH_PY419_DataValidCheck() == false)
                        //{
                            BubbleEvent = false;
                        //}
                    }
                    if (pVal.ItemUID == "Btn_ret")
                    {
                        PH_PY419_MTX01();
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        PH_PY419_SAVE();

                    }
                    if (pVal.ItemUID == "Btn_del")
                    {
                        PH_PY419_Delete();
                        PH_PY419_FormItemEnabled();
                    }
                    //                If oForm.Mode = fm_FIND_MODE Then
                    //                    If pVal.ItemUID = "Btn01" Then
                    //                        Sbo_Application.ActivateMenuItem ("7425")
                    //                        BubbleEvent = False
                    //                    End If
                    //
                    //                End If
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY419_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY419_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY419_FormItemEnabled();
                                }
                            }
                            break;
                            //
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
        
        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
            string sQry = string.Empty;
            try
            {
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {

                    }

                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {

                            case "MSTCOD":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

                                sQry = "Select Code,";
                                sQry = sQry + " FullName = U_FullName ";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
                                break;

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
            string sQry = string.Empty;
            try
            {
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        //Call oMat1.SelectRow(pVal.Row, True, False)
                                        oForm.Items.Item("Year").Specific.VALUE = oDS_PH_PY419.Columns.Item("Year").Cells.Item(pVal.Row).Value;
                                        oForm.Items.Item("MSTCOD").Specific.VALUE = oDS_PH_PY419.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                                        oForm.Items.Item("FullName").Specific.VALUE = oDS_PH_PY419.Columns.Item("FullName").Cells.Item(pVal.Row).Value;

                                        break;
                                }

                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
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
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataFind : 자료 삭제
        /// </summary>
        private void PH_PY419_Delete()
        {
            short cnt = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string FullName = string.Empty;
            string MSTCOD = string.Empty;
            string Year = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                //FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();

                sQry = " Select Count(*) as cnt From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(cnt).Value > 0)
                {

                    if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.VALUE))
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                    {
                        ErrNum = 2;
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE))
                    {
                        ErrNum = 3;
                        throw new Exception();
                    }
                    if (PSH_Globals.SBO_Application.MessageBox(" 선택한사원('" + oForm.Items.Item("FullName").Specific.Value.ToString().Trim() + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                    {
                        sQry = "Delete From [p_seoyst] Where saup = '" + CLTCOD + "' AND  yyyy = '" + Year + "' And sabun = '" + MSTCOD + "' ";
                        oRecordSet.DoQuery(sQry);
                    }
                }
                PH_PY419_MTX01();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_Delete_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }


        //private void PH_PY419_Delete()
        //{
        //    //선택된 자료 삭제

        //    string CLTCOD = null;
        //    string MSTCOD = null;
        //    UPGRADE_NOTE: YEAR이(가) Year(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //    string Year = null;
        //    string FullName = null;


        //    short i = 0;
        //    short cnt = 0;

        //    string sQry = null;

        //    SAPbobsCOM.Recordset oRecordSet = null;

        //     ERROR: Not supported in C#: OnErrorStatement


        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



        //    oForm.Freeze(true);

        //    UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //    UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Year = oForm.Items.Item("Year").Specific.VALUE;
        //    UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

        //    sQry = " Select Count(*) From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
        //    oRecordSet.DoQuery(sQry);

        //    UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    cnt = oRecordSet.Fields.Item(0).Value;
        //    if (cnt > 0)
        //    {

        //        if (string.IsNullOrEmpty(Strings.Trim(Year)))
        //        {
        //            MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //            goto PH_PY419_Delete_Exit;
        //        }

        //        if (string.IsNullOrEmpty(Strings.Trim(CLTCOD)))
        //        {
        //            MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //            goto PH_PY419_Delete_Exit;
        //        }
        //        if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
        //        {
        //            MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
        //            goto PH_PY419_Delete_Exit;
        //        }




        //        if (MDC_Globals.Sbo_Application.MessageBox(" 선택한사원('" + FullName + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
        //        {
        //            sQry = "Delete From [p_seoyst] Where saup = '" + CLTCOD + "' AND  yyyy = '" + Year + "' And sabun = '" + MSTCOD + "' ";
        //            oRecordSet.DoQuery(sQry);
        //        }
        //    }


        //    oForm.Freeze(false);


        //    PH_PY419_MTX01();

        //    UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;


        //    return;
        //PH_PY419_Delete_Exit:
        //    UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;

        //    oForm.Freeze(false);
        //    return;
        //PH_PY419_Delete_Error:
        //    UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;

        //    oForm.Freeze(false);
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}

        /// <summary>
        /// DataFind : 자료 삭제
        /// </summary>
        private void PH_PY419_MTX01()
        {
            short ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string FullName = string.Empty;
            string MSTCOD = string.Empty;
            string Year = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Year))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(CLTCOD))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                sQry = "EXEC PH_PY419_01 '" + CLTCOD + "', '" + Year + "'";

                oDS_PH_PY419.ExecuteQuery(sQry);

            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년도가 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_Delete_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }


        //private void PH_PY419_MTX01()
        //{

        //    ////메트릭스에 데이터 로드

        //    int i = 0;
        //    string sQry = null;
        //    int iRow = 0;

        //    string Param01 = null;
        //    string Param02 = null;

        //    SAPbobsCOM.Recordset oRecordSet = null;

        //    // ERROR: Not supported in C#: OnErrorStatement


        //    oForm.Freeze(true);
        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param02 = oForm.Items.Item("Year").Specific.VALUE;

        //    if (string.IsNullOrEmpty(Strings.Trim(Param01)))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //        goto PH_PY419_MTX01_Exit;
        //    }

        //    if (string.IsNullOrEmpty(Strings.Trim(Param02)))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //        goto PH_PY419_MTX01_Exit;
        //    }



        //    sQry = "EXEC PH_PY419_01 '" + Param01 + "', '" + Param02 + "'";

        //    oDS_PH_PY419A.ExecuteQuery(sQry);



        //    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

        //    PH_PY419_TitleSetting(ref iRow);

        //    oForm.Update();

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);
        //    return;
        //PH_PY419_MTX01_Exit:
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);
        //    return;
        //PH_PY419_MTX01_Error:
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}

        /// <summary>
        /// DataFind : 자료 조회
        /// </summary>
        private void PH_PY419_SAVE()
        {
            short ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string FullName = string.Empty;
            string MSTCOD = string.Empty;
            string Year = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.VALUE))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE))
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                sQry = " Select Count(*) From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {

                }
                else
                {

                    ////신규
                    sQry = "INSERT INTO [p_seoyst]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "kname";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + CLTCOD + "',";
                    sQry = sQry + "'" + Year + "',";
                    sQry = sQry + "'" + MSTCOD + "',";
                    sQry = sQry + "'" + FullName + "'";
                    sQry = sQry + ")";

                    oRecordSet.DoQuery(sQry);
                }
                PH_PY419_MTX01();
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년도가 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번이 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY419_SAVE_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        //private void PH_PY419_SAVE()
        //{

        //    ////데이타 저장

        //    int i = 0;
        //    string sQry = null;

        //    //UPGRADE_NOTE: YEAR이(가) Year(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //    string FullName = null;
        //    string CLTCOD = null;
        //    string MSTCOD = null;
        //    string Year = null;

        //    SAPbobsCOM.Recordset oRecordSet = null;

        //    oForm.Freeze(true);
        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //    Year = oForm.Items.Item("Year").Specific.VALUE;
        //    MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
        //    FullName = oForm.Items.Item("FullName").Specific.VALUE;

        //    if (string.IsNullOrEmpty(Strings.Trim(Year)))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //        goto PH_PY419_SAVE_Exit;
        //    }

        //    if (string.IsNullOrEmpty(Strings.Trim(CLTCOD)))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //        goto PH_PY419_SAVE_Exit;
        //    }
        //    if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
        //        goto PH_PY419_SAVE_Exit;
        //    }

        //    sQry = " Select Count(*) From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
        //    oRecordSet.DoQuery(sQry);

        //    if (oRecordSet.Fields.Item(0).Value > 0)
        //    {
        //    }
        //    else
        //    {

        //        ////신규
        //        sQry = "INSERT INTO [p_seoyst]";
        //        sQry = sQry + " (";
        //        sQry = sQry + "saup,";
        //        sQry = sQry + "yyyy,";
        //        sQry = sQry + "sabun,";
        //        sQry = sQry + "kname";
        //        sQry = sQry + " ) ";
        //        sQry = sQry + "VALUES(";

        //        sQry = sQry + "'" + CLTCOD + "',";
        //        sQry = sQry + "'" + Year + "',";
        //        sQry = sQry + "'" + MSTCOD + "',";
        //        sQry = sQry + "'" + FullName + "'";
        //        sQry = sQry + ")";

        //        oRecordSet.DoQuery(sQry);
        //    }


        //    PH_PY419_FormItemEnabled();


        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);

        //    PH_PY419_MTX01();

        //    return;
        //PH_PY419_SAVE_Exit:

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);

        //    return;
        //PH_PY419_SAVE_Error:
        //    oForm.Freeze(false);

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}

        //public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    string sQry = null;
        //    int i = 0;
        //    string tSex = null;
        //    string tBrith = null;
        //    //UPGRADE_NOTE: Day이(가) Day_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //    string Day_Renamed = null;
        //    string ActCode = null;
        //    string CLTCOD = null;
        //    string MSTCOD = null;

        //    SAPbouiCOM.ComboBox oCombo = null;
        //    SAPbouiCOM.Column oColumn = null;
        //    SAPbouiCOM.Columns oColumns = null;
        //    SAPbobsCOM.Recordset oRecordSet = null;

        //    // ERROR: Not supported in C#: OnErrorStatement


        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    switch (pVal.EventType)
        //    {
        //        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //            ////1

        //            if (pVal.BeforeAction == true)
        //            {
        //                if (pVal.ItemUID == "1")
        //                {
        //                    if (PH_PY419_DataValidCheck() == false)
        //                    {
        //                        BubbleEvent = false;
        //                    }
        //                }

        //                if (pVal.ItemUID == "Btn_ret")
        //                {
        //                    PH_PY419_MTX01();
        //                }



        //                if (pVal.ItemUID == "Btn01")
        //                {
        //                    PH_PY419_SAVE();

        //                }


        //                if (pVal.ItemUID == "Btn_del")
        //                {
        //                    PH_PY419_Delete();
        //                    PH_PY419_FormItemEnabled();
        //                }
        //                //                If oForm.Mode = fm_FIND_MODE Then
        //                //                    If pVal.ItemUID = "Btn01" Then
        //                //                        Sbo_Application.ActivateMenuItem ("7425")
        //                //                        BubbleEvent = False
        //                //                    End If
        //                //
        //                //                End If
        //            }
        //            else if (pVal.BeforeAction == false)
        //            {
        //                switch (pVal.ItemUID)
        //                {
        //                    case "1":
        //                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                        {
        //                            if (pVal.ActionSuccess == true)
        //                            {
        //                                PH_PY419_FormItemEnabled();
        //                            }
        //                        }
        //                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //                        {
        //                            if (pVal.ActionSuccess == true)
        //                            {
        //                                PH_PY419_FormItemEnabled();
        //                            }
        //                        }
        //                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                        {
        //                            if (pVal.ActionSuccess == true)
        //                            {
        //                                PH_PY419_FormItemEnabled();
        //                            }
        //                        }
        //                        break;
        //                        //
        //                }
        //            }
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //            ////2
        //            if (pVal.BeforeAction == true)
        //            {
        //                if (pVal.CharPressed == 9)
        //                {
        //                    if (pVal.ItemUID == "MSTCOD")
        //                    {
        //                        //UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE))
        //                        {
        //                            MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
        //                            BubbleEvent = false;
        //                        }
        //                    }
        //                }
        //            }
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //            ////3
        //            switch (pVal.ItemUID)
        //            {
        //                case "Mat01":
        //                    if (pVal.Row > 0)
        //                    {
        //                        oLastItemUID = pVal.ItemUID;
        //                        oLastColUID = pVal.ColUID;
        //                        oLastColRow = pVal.Row;
        //                    }
        //                    break;
        //                default:
        //                    oLastItemUID = pVal.ItemUID;
        //                    oLastColUID = "";
        //                    oLastColRow = 0;
        //                    break;
        //            }
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //            ////4
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //            ////5
        //            oForm.Freeze(true);
        //            if (pVal.BeforeAction == true)
        //            {

        //            }
        //            else if (pVal.BeforeAction == false)
        //            {
        //                if (pVal.ItemChanged == true)
        //                {
        //                    ////사업장(헤더)
        //                    if (pVal.ItemUID == "SCLTCOD")
        //                    {

        //                    }

        //                }
        //            }

        //            oForm.Freeze(false);
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_CLICK:
        //            ////6
        //            oForm.Freeze(true);
        //            if (pVal.BeforeAction == true)
        //            {
        //                switch (pVal.ItemUID)
        //                {
        //                    case "Grid01":
        //                        if (pVal.Row >= 0)
        //                        {
        //                            switch (pVal.ItemUID)
        //                            {
        //                                case "Grid01":
        //                                    //Call oMat1.SelectRow(pVal.Row, True, False)
        //                                    PH_PY419_MTX02(pVal.ItemUID, ref pVal.Row, ref pVal.ColUID);
        //                                    break;
        //                            }

        //                        }
        //                        break;
        //                }

        //                switch (pVal.ItemUID)
        //                {
        //                    case "Grid01":
        //                        if (pVal.Row > 0)
        //                        {
        //                            oLastItemUID = pVal.ItemUID;
        //                            oLastColUID = pVal.ColUID;
        //                            oLastColRow = pVal.Row;
        //                        }
        //                        break;
        //                    default:
        //                        oLastItemUID = pVal.ItemUID;
        //                        oLastColUID = "";
        //                        oLastColRow = 0;
        //                        break;
        //                }
        //            }
        //            else if (pVal.BeforeAction == false)
        //            {

        //            }
        //            oForm.Freeze(false);
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //            ////7
        //            oForm.Freeze(true);
        //            if (pVal.BeforeAction == true)
        //            {
        //            }
        //            else
        //            {

        //            }
        //            oForm.Freeze(false);
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //            ////8
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
        //            ////9
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //            ////10
        //            //            Call oForm.Freeze(True)
        //            if (pVal.BeforeAction == true)
        //            {
        //                if (pVal.ItemChanged == true)
        //                {

        //                }

        //            }
        //            else if (pVal.BeforeAction == false)
        //            {
        //                if (pVal.ItemChanged == true)
        //                {
        //                    switch (pVal.ItemUID)
        //                    {

        //                        case "MSTCOD":
        //                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

        //                            sQry = "Select Code,";
        //                            sQry = sQry + " FullName = U_FullName ";
        //                            sQry = sQry + " From [@PH_PY001A]";
        //                            sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
        //                            sQry = sQry + " and Code = '" + MSTCOD + "'";

        //                            oRecordSet.DoQuery(sQry);

        //                            //UPGRADE_WARNING: oForm.Items(FullName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            //UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
        //                            break;

        //                    }

        //                }
        //            }
        //            break;
        //        //            Call oForm.Freeze(False)
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //            ////11
        //            if (pVal.BeforeAction == true)
        //            {
        //            }
        //            else if (pVal.BeforeAction == false)
        //            {
        //                //                oMat1.LoadFromDataSource
        //                //                Call PH_PY419_AddMatrixRow

        //            }
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
        //            ////12
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
        //            ////16
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //            ////17
        //            if (pVal.BeforeAction == true)
        //            {
        //            }
        //            else if (pVal.BeforeAction == false)
        //            {
        //                SubMain.RemoveForms(oFormUniqueID);
        //                //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oForm = null;
        //                //UPGRADE_NOTE: oDS_PH_PY419A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                oDS_PH_PY419A = null;

        //                //                Set oMat1 = Nothing
        //            }
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //            ////18
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //            ////19
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
        //            ////20
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //            ////21
        //            if (pVal.BeforeAction == true)
        //            {

        //            }
        //            else if (pVal.BeforeAction == false)
        //            {

        //            }
        //            break;
        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
        //            ////22
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
        //            ////23
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //            ////27
        //            if (pVal.BeforeAction == true)
        //            {

        //            }
        //            else if (pVal.Before_Action == false)
        //            {
        //                //                If pVal.ItemUID = "Code" Then
        //                //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY419A", "Code")
        //                //                End If
        //            }
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
        //            ////37
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
        //            ////38
        //            break;

        //        //----------------------------------------------------------
        //        case SAPbouiCOM.BoEventTypes.et_Drag:
        //            ////39
        //            break;


        //    }

        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;

        //    return;
        //Raise_FormItemEvent_Error:
        //    ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //    oForm.Freeze((false));
        //    //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oCombo = null;
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}


        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //		{
        //			int i = 0;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm.Freeze(true);

        //			if ((pVal.BeforeAction == true)) {
        //				switch (pVal.MenuUID) {
        //					case "1283":
        //						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //						break;
        //					case "1284":
        //						break;
        //					case "1286":
        //						break;
        //					case "1293":
        //						break;
        //					case "1281":
        //						break;
        //					case "1282":
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						break;
        //					//                Call PH_PY419_FormItemEnabled
        //				}
        //			} else if ((pVal.BeforeAction == false)) {
        //				switch (pVal.MenuUID) {
        //					case "1283":
        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //						PH_PY419_FormItemEnabled();
        //						break;
        //					//                Call PH_PY419_AddMatrixRow
        //					case "1284":
        //						break;
        //					case "1286":
        //						break;
        //					//            Case "1293":
        //					//                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
        //					case "1281":
        //						////문서찾기
        //						PH_PY419_FormItemEnabled();
        //						//                Call PH_PY419_AddMatrixRow
        //						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						break;
        //					case "1282":
        //						////문서추가
        //						PH_PY419_FormItemEnabled();
        //						break;
        //					//                Call PH_PY419_AddMatrixRow
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						PH_PY419_FormItemEnabled();
        //						break;
        //					case "1293":
        //						//// 행삭제
        //						break;

        //				}
        //			}
        //			oForm.Freeze(false);
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			int i = 0;
        //			string sQry = null;
        //			SAPbouiCOM.ComboBox oCombo = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;


        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}

        //			}
        //			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oCombo = null;
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			Raise_FormDataEvent_Error:

        //			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oCombo = null;
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //		}

        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pVal.BeforeAction == true) {
        //			} else if (pVal.BeforeAction == false) {
        //			}
        //			switch (pVal.ItemUID) {
        //				case "Mat01":
        //					if (pVal.Row > 0) {
        //						oLastItemUID = pVal.ItemUID;
        //						oLastColUID = pVal.ColUID;
        //						oLastColRow = pVal.Row;
        //					}
        //					break;
        //				default:
        //					oLastItemUID = pVal.ItemUID;
        //					oLastColUID = "";
        //					oLastColRow = 0;
        //					break;
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		public void PH_PY419_FormClear()
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			string DocEntry = null;
        //			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY419'", ref "");
        //			if (Convert.ToDouble(DocEntry) == 0) {
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
        //			} else {
        //				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //			}
        //			return;
        //			PH_PY419_FormClear_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		public bool PH_PY419_DataValidCheck()
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			functionReturnValue = false;
        //			int i = 0;
        //			int j = 0;

        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			return functionReturnValue;


        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			PH_PY419_DataValidCheck_Error:


        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			functionReturnValue = false;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}

        //		private void PH_PY419_MTX01()
        //		{

        //			////메트릭스에 데이터 로드

        //			int i = 0;
        //			string sQry = null;
        //			int iRow = 0;

        //			string Param01 = null;
        //			string Param02 = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Param02 = oForm.Items.Item("Year").Specific.VALUE;

        //			if (string.IsNullOrEmpty(Strings.Trim(Param01))) {
        //				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //				goto PH_PY419_MTX01_Exit;
        //			}

        //			if (string.IsNullOrEmpty(Strings.Trim(Param02))) {
        //				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //				goto PH_PY419_MTX01_Exit;
        //			}



        //			sQry = "EXEC PH_PY419_01 '" + Param01 + "', '" + Param02 + "'";

        //			oDS_PH_PY419A.ExecuteQuery(sQry);



        //			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

        //			PH_PY419_TitleSetting(ref iRow);

        //			oForm.Update();

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			oForm.Freeze(false);
        //			return;
        //			PH_PY419_MTX01_Exit:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			oForm.Freeze(false);
        //			return;
        //			PH_PY419_MTX01_Error:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        //private void PH_PY419_MTX02(string oUID, ref int oRow = 0, ref string oCol = "")
        //{


        //    ////그리드 자료를 head에 로드

        //    int i = 0;
        //    string sQry = null;
        //    int sRow = 0;

        //    string Param01 = null;
        //    string Param02 = null;
        //    string Param03 = null;

        //    SAPbouiCOM.ComboBox oCombo = null;
        //    SAPbobsCOM.Recordset oRecordSet = null;

        //    // ERROR: Not supported in C#: OnErrorStatement


        //    oForm.Freeze(true);
        //    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    sRow = oRow;

        //   oForm.Items.Item("Year").Specific.VALUE = oDS_PH_PY419A.Columns.Item("Year").Cells.Item(oRow).Value;
        //   oForm.Items.Item("MSTCOD").Specific.VALUE = oDS_PH_PY419A.Columns.Item("MSTCOD").Cells.Item(oRow).Value;
        //   oForm.Items.Item("FullName").Specific.VALUE = oDS_PH_PY419A.Columns.Item("FullName").Cells.Item(oRow).Value;


        //    oForm.Update();

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);
        //    return;
        //PH_PY419_MTX02_Exit:
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);

        //    return;
        //PH_PY419_MTX02_Error:
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    oForm.Freeze(false);
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}

        //		public bool PH_PY419_Validate(string ValidateType)
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			functionReturnValue = true;
        //			object i = null;
        //			int j = 0;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY419A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY419A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				functionReturnValue = false;
        //				goto PH_PY419_Validate_Exit;
        //			}
        //			//
        //			if (ValidateType == "수정") {

        //			} else if (ValidateType == "행삭제") {

        //			} else if (ValidateType == "취소") {

        //			}
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY419_Validate_Exit:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return functionReturnValue;
        //			PH_PY419_Validate_Error:
        //			functionReturnValue = false;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			return functionReturnValue;
        //		}

        //////행삭제 (FormUID, pVal, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
        //		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent, ref SAPbouiCOM.Matrix oMat, ref SAPbouiCOM.DBDataSource DBData, ref string CheckField)
        //		{

        //			int i = 0;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if ((oLastColRow > 0)) {
        //				if (pVal.BeforeAction == true) {

        //				} else if (pVal.BeforeAction == false) {
        //					if (oMat.RowCount != oMat.VisualRowCount) {
        //						oMat.FlushToDataSource();

        //						while ((i <= DBData.Size - 1)) {
        //							if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i))) {
        //								DBData.RemoveRecord((i));
        //								i = 0;
        //							} else {
        //								i = i + 1;
        //							}
        //						}

        //						for (i = 0; i <= DBData.Size; i++) {
        //							DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //						}

        //						oMat.LoadFromDataSource();
        //					}
        //				}
        //			}
        //			return;
        //			Raise_EVENT_ROW_DELETE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void PH_PY419_SAVE()
        //		{

        //			////데이타 저장

        //			int i = 0;
        //			string sQry = null;

        //			//UPGRADE_NOTE: YEAR이(가) Year(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //			string FullName = null;
        //			string CLTCOD = null;
        //			string MSTCOD = null;
        //			string Year = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);
        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Year = oForm.Items.Item("Year").Specific.VALUE;
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			FullName = oForm.Items.Item("FullName").Specific.VALUE;

        //			if (string.IsNullOrEmpty(Strings.Trim(Year))) {
        //				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //				goto PH_PY419_SAVE_Exit;
        //			}

        //			if (string.IsNullOrEmpty(Strings.Trim(CLTCOD))) {
        //				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //				goto PH_PY419_SAVE_Exit;
        //			}
        //			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
        //				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
        //				goto PH_PY419_SAVE_Exit;
        //			}

        //			sQry = " Select Count(*) From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
        //			oRecordSet.DoQuery(sQry);

        //			if (oRecordSet.Fields.Item(0).Value > 0) {
        //				////갱신

        //				//        sQry = "Update [p_sbservcomp] set "
        //				//        sQry = sQry + "entno1 = '" & entno1 & "',"
        //				//        sQry = sQry + "servcomp1 = '" & servcomp1 & "',"
        //				//        sQry = sQry + "symd1 = '" & symd1 & "',"
        //				//        sQry = sQry + "eymd1 = '" & eymd1 & "',"
        //				//        sQry = sQry + "payrtot1 = " & payrtot1 & ","
        //				//        sQry = sQry + "bnstot1 = " & bnstot1 & ","
        //				//        sQry = sQry + "fwork1 = " & fwork1 & ","
        //				//        sQry = sQry + "ndtalw1 = " & ndtalw1 & ","
        //				//        sQry = sQry + "etcntax1 = " & etcntax1 & ","
        //				//        sQry = sQry + "lnchalw1 = " & lnchalw1 & ","
        //				//        sQry = sQry + "ftaxamt1 = " & ftaxamt1 & ","
        //				//        sQry = sQry + "savtaxddc1 = " & savtaxddc1 & ","
        //				//        sQry = sQry + "incmtax1 = " & incmtax1 & ","
        //				//        sQry = sQry + "fvsptax1 = " & fvsptax1 & ","
        //				//        sQry = sQry + "residtax1 = " & residtax1 & ","
        //				//        sQry = sQry + "medcinsr1 = " & medcinsr1 & ","
        //				//        sQry = sQry + "asopinsr1 = " & asopinsr1 & ","
        //				//        sQry = sQry + "annuboamt1 =" & annuboamt1 & ","
        //				//        sQry = sQry + "entno2 = '" & entno2 & "',"
        //				//        sQry = sQry + "servcomp2 = '" & servcomp2 & "',"
        //				//        sQry = sQry + "symd2 = '" & symd2 & "',"
        //				//        sQry = sQry + "eymd2 = '" & eymd2 & "',"
        //				//        sQry = sQry + "payrtot2 = " & payrtot2 & ","
        //				//        sQry = sQry + "bnstot2= " & bnstot2 & ","
        //				//        sQry = sQry + "fwork2 = " & fwork2 & ","
        //				//        sQry = sQry + "ndtalw2 = " & ndtalw2 & ","
        //				//        sQry = sQry + "etcntax2 = " & etcntax2 & ","
        //				//        sQry = sQry + "lnchalw2 = " & lnchalw2 & ","
        //				//        sQry = sQry + "ftaxamt2 = " & ftaxamt2 & ","
        //				//        sQry = sQry + "savtaxddc2 = " & savtaxddc2 & ","
        //				//        sQry = sQry + "indmtax2 = " & indmtax2 & ","
        //				//        sQry = sQry + "fvsptax2 = " & fvsptax2 & ","
        //				//        sQry = sQry + "residtax2 = " & residtax2 & ","
        //				//        sQry = sQry + "medcinsr2 = " & medcinsr2 & ","
        //				//        sQry = sQry + "asopinsr2 = " & asopinsr2 & ","
        //				//        sQry = sQry + "annuboamt2 =" & annuboamt2
        //				//
        //				//        sQry = sQry + " Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
        //				//
        //				//        oRecordSet.DoQuery sQry

        //			} else {

        //				////신규
        //				sQry = "INSERT INTO [p_seoyst]";
        //				sQry = sQry + " (";
        //				sQry = sQry + "saup,";
        //				sQry = sQry + "yyyy,";
        //				sQry = sQry + "sabun,";
        //				sQry = sQry + "kname";
        //				sQry = sQry + " ) ";
        //				sQry = sQry + "VALUES(";

        //				sQry = sQry + "'" + CLTCOD + "',";
        //				sQry = sQry + "'" + Year + "',";
        //				sQry = sQry + "'" + MSTCOD + "',";
        //				sQry = sQry + "'" + FullName + "'";
        //				sQry = sQry + ")";

        //				oRecordSet.DoQuery(sQry);
        //			}


        //			PH_PY419_FormItemEnabled();


        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			oForm.Freeze(false);

        //			PH_PY419_MTX01();

        //			return;
        //			PH_PY419_SAVE_Exit:

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			oForm.Freeze(false);

        //			return;
        //			PH_PY419_SAVE_Error:
        //			oForm.Freeze(false);

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void PH_PY419_Delete()
        //		{
        //			////선택된 자료 삭제

        //			string CLTCOD = null;
        //			string MSTCOD = null;
        //			//UPGRADE_NOTE: YEAR이(가) Year(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //			string Year = null;
        //			string FullName = null;


        //			short i = 0;
        //			short cnt = 0;

        //			string sQry = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



        //			oForm.Freeze(true);

        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Year = oForm.Items.Item("Year").Specific.VALUE;
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

        //			sQry = " Select Count(*) From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
        //			oRecordSet.DoQuery(sQry);

        //			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			cnt = oRecordSet.Fields.Item(0).Value;
        //			if (cnt > 0) {

        //				if (string.IsNullOrEmpty(Strings.Trim(Year))) {
        //					MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
        //					goto PH_PY419_Delete_Exit;
        //				}

        //				if (string.IsNullOrEmpty(Strings.Trim(CLTCOD))) {
        //					MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
        //					goto PH_PY419_Delete_Exit;
        //				}
        //				if (string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
        //					MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
        //					goto PH_PY419_Delete_Exit;
        //				}




        //				if (MDC_Globals.Sbo_Application.MessageBox(" 선택한사원('" + FullName + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1")) {
        //					sQry = "Delete From [p_seoyst] Where saup = '" + CLTCOD + "' AND  yyyy = '" + Year + "' And sabun = '" + MSTCOD + "' ";
        //					oRecordSet.DoQuery(sQry);
        //				}
        //			}


        //			oForm.Freeze(false);


        //			PH_PY419_MTX01();

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;


        //			return;
        //			PH_PY419_Delete_Exit:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;

        //			oForm.Freeze(false);
        //			return;
        //			PH_PY419_Delete_Error:
        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void PH_PY419_TitleSetting(ref int iRow)
        //		{
        //			int i = 0;
        //			int j = 0;
        //			string sQry = null;

        //			string[] COLNAM = new string[3];

        //			SAPbouiCOM.EditTextColumn oColumn = null;
        //			SAPbouiCOM.ComboBoxColumn oComboCol = null;

        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			oForm.Freeze(true);

        //			COLNAM[0] = "년도";
        //			COLNAM[1] = "사번";
        //			COLNAM[2] = "성명";


        //			for (i = 0; i <= Information.UBound(COLNAM); i++) {
        //				oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
        //				oGrid1.Columns.Item(i).Editable = false;

        //				//    oGrid1.Columns.Item(i).RightJustified = True

        //			}

        //			oGrid1.AutoResizeColumns();

        //			oForm.Freeze(false);

        //			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oColumn = null;

        //			return;
        //			Error_Message:

        //			oForm.Freeze(false);
        //			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oColumn = null;
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY419_TitleSetting Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
    }
}
