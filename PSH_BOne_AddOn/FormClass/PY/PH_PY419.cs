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
        public override void LoadForm(string oFormDocEntry01)
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

                

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //    oForm.DataBrowser.BrowseBy = "Code"

                oForm.Freeze(true);
                PH_PY419_CreateItems();
                PH_PY419_EnableMenus();
                PH_PY419_SetDocument(oFormDocEntry01);
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

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY419_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY419_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY419_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
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
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.Value);
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                                sQry = "Select Code,";
                                sQry = sQry + " FullName = U_FullName ";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
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
                                        oForm.Items.Item("Year").Specific.Value = oDS_PH_PY419.Columns.Item("Year").Cells.Item(pVal.Row).Value;
                                        oForm.Items.Item("MSTCOD").Specific.Value = oDS_PH_PY419.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                                        oForm.Items.Item("FullName").Specific.Value = oDS_PH_PY419.Columns.Item("FullName").Cells.Item(pVal.Row).Value;

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
            string sQry;
            string CLTCOD;
            string MSTCOD;
            string Year;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                sQry = " Select Count(*) as cnt From [p_seoyst] Where saup = '" + CLTCOD + "' And yyyy = '" + Year + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet.Fields.Item(cnt).Value > 0))
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value))
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                    {
                        ErrNum = 2;
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
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
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년도를 입력하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장을 입력하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번을 입력하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

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

                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
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
    }
}
