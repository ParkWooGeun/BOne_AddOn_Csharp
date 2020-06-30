using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사용자 권한 등록
    /// </summary>
    internal class PH_PY998 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY998A;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY998.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY998_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY998");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                //PH_PY998_CreateItems();
                //PH_PY998_FormItemEnabled();
                //PH_PY998_EnableMenus();
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
                //oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY998_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY998");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY998");
                oDS_PH_PY998A = oForm.DataSources.DataTables.Item("PH_PY998");

                // 구분
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("B", "기본");
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("H", "인사");
                oForm.Items.Item("pGubun").DisplayDesc = true;

                // 폴더/화면구분
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("F", "폴더");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("S", "화면");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("C", "복제");
                oForm.Items.Item("pFSGubun").DisplayDesc = true;

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A006'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Modual").Specific, "Y");
                oForm.Items.Item("Modual").DisplayDesc = true;

                // Position
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A005'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Position").Specific, "Y");
                oForm.Items.Item("Position").DisplayDesc = true;

                // Sub1
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub1").Specific, "Y");
                oForm.Items.Item("Sub1").DisplayDesc = true;

                // Sub2
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub2").Specific, "Y");
                oForm.Items.Item("Sub2").DisplayDesc = true;

                // Sub3
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub3").Specific, "Y");
                oForm.Items.Item("Sub3").DisplayDesc = true;

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("No").Specific, "Y");
                oForm.Items.Item("No").DisplayDesc = true;

                // Level
                oForm.Items.Item("Level").Specific.ValidValues.Add("0", "0");
                oForm.Items.Item("Level").Specific.ValidValues.Add("1", "1");
                oForm.Items.Item("Level").Specific.ValidValues.Add("2", "2");
                oForm.Items.Item("Level").DisplayDesc = true;

                // FatherId
                sQry = "select  distinct t.a,t.b   from (select distinct UniqueID as a , UniqueID as b from Authority_Folder union all select distinct FatherID as a , FatherID as b from Authority_Folder) t";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FatherID").Specific, "Y");
                oForm.Items.Item("FatherID").DisplayDesc = true;

                // String
                oForm.DataSources.UserDataSources.Add("Strings", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("Strings").Specific.DataBind.SetBound(true, "", "Strings");

                // UniqueID
                oForm.DataSources.UserDataSources.Add("UniqueID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("UniqueID").Specific.DataBind.SetBound(true, "", "UniqueID");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY998_EnableMenus
        /// </summary>
        private void PH_PY998_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false);                // 제거
                oForm.EnableMenu("1284", false);                // 취소
                oForm.EnableMenu("1293", false);                // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY998_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                {
                    oForm.Items.Item("pUserID").Enabled = false;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = true;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;

                    oForm.Items.Item("BtnSearch").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;

                }

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "S"))
                {
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = true;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = true;

                    oForm.Items.Item("BtnSearch").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;

                }
                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "C"))
                {
                    oForm.Items.Item("CPUserID").Enabled = true;
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("pGubun").Enabled = false;
                    oForm.Items.Item("Modual").Enabled = false;
                    oForm.Items.Item("Sub1").Enabled = false;
                    oForm.Items.Item("Sub2").Enabled = false;
                    oForm.Items.Item("Sub3").Enabled = false;
                    oForm.Items.Item("No").Enabled = false;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = false;
                    oForm.Items.Item("FatherID").Enabled = false;
                    oForm.Items.Item("Position").Enabled = false;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;

                    oForm.Items.Item("BtnSearch").Enabled = false;
                    oForm.Items.Item("Bt_Copy").Enabled = true;
                }

                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.EnableMenu("1281", false);      // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.EnableMenu("1281", false);      // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.EnableMenu("1281", true);       // 문서찾기
                    oForm.EnableMenu("1282", true);       // 문서추가
                }
                // Key set
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY998_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY998A);
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
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
            }
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
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
                            PH_PY998_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY998_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "1293": // 행삭제
                            break;
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
                    }

                    if (pVal.ItemUID == "BtnSearch")
                    {
                        PH_PY998_MTX01();
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        PH_PY998_SAVE(pVal.Row);
                    }

                    if (pVal.ItemUID == "Btn_del")
                    {
                        PH_PY998_Delete();
                    }
                    if (pVal.ItemUID == "Bt_Copy")
                    {
                        PH_PY998_Copy();
                    }
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
                                    PH_PY998_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY998_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY998_FormItemEnabled();
                                }
                            }
                            break;
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        // 사업장(헤더)
                        switch (pVal.ItemUID)
                        {
                            case "Modual":
                            case "No":
                            case "Sub1":
                            case "Sub2":
                            case "Sub3":
                                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                                {
                                    oForm.DataSources.UserDataSources.Item("UniqueId").Value = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                    oForm.Items.Item("Sequence").Specific.VALUE = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                }
                                oForm.Items.Item("Sequence").Specific.VALUE = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim() + oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim() + oForm.Items.Item("No").Specific.VALUE.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                                break;
                        }
                    }
                }
                PH_PY998_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
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
                                        PH_PY998_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
        /// PH_PY998_MTX01
        /// </summary>
        private void PH_PY998_MTX01()
        {
            //int iRow = 0;
            //int ErrNum = 0;
            //string sQry = string.Empty;
            //string Param01 = string.Empty;
            //string Param02 = string.Empty;
            //string Param03 = string.Empty;


            SAPbobsCOM.SBObob oSBObob = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                /*
                    1. 현재 사용자 ID 조회, 재직중인 사원(OUSR User_Code 조회)
                    2. 1의 카운트만큼 루프
                    3. 조회하고자 하는 권한(1:모든 권한, 2:읽기 전용, 3:권한 없음)을 가진 사용자 ID를 저장(DataRow?)
                    4. 저장된 DataRow의 카운트만큼 루프
                        4-1. matrix의 각 필드에 매칭 데이터 출력
                */



                //int userListCount = oSBObob.GetUserList().RecordCount;

                //for (int loopCount = 0; loopCount <= userListCount; loopCount++)
                //{
                //    PSH_Globals.SBO_Application.MessageBox(oSBObob..GetUserList(loopCount).ToString());
                //}


                //oRecordSet = oSBObob.GetSystemPermission("manager", "142");

                //Debug.WriteLine(oRecordSet.Fields.Item(0).Value())

                PSH_Globals.SBO_Application.MessageBox(oRecordSet.Fields.Item(0).Value);

                //Param01 = oForm.Items.Item("pGubun").Specific.VALUE.ToString().Trim();
                //Param02 = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                //Param03 = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();

                //if (string.IsNullOrEmpty(Param01.ToString().Trim()))
                //{
                //    ErrNum = 1;
                //    throw new Exception();
                //}

                //if (Param02 == "S")
                //{
                //    if (string.IsNullOrEmpty(Param03.ToString().Trim()))
                //    {
                //        ErrNum = 2;
                //        throw new Exception();
                //    }
                //}

                //sQry = "EXEC PH_PY998_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                //oDS_PH_PY998A.ExecuteQuery(sQry);
                //iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
            }
            catch (Exception ex)
            {
                //if (ErrNum == 1)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText("구분이 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
                //else if (ErrNum == 2)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText("USERID가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
                //else
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}

                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY998_MTX02
        /// </summary>
        private void PH_PY998_MTX02(string oUID, int oRow = 0, string oCol = "")
        {
            int sRow = 0;
            int ErrNum = 0;
            string sQry = string.Empty;
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                sRow = oRow;
                Param01 = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                Param02 = oDS_PH_PY998A.Columns.Item("UniqueID").Cells.Item(oRow).Value;
                if (Param01 != "F")
                {
                    Param03 = oDS_PH_PY998A.Columns.Item("UserID").Cells.Item(oRow).Value;
                }
                sQry = "EXEC PH_PY998_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                if ((oRecordSet.RecordCount == 0))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                // Screen일때 UserID를 가져옴.
                if (Param01 != "F")
                {
                    // oForm.Items.Item("UserID").Specific.Select(oRecordSet.Fields.Item("UserID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("UserID").Specific.value = oRecordSet.Fields.Item("UserID").Value;
                }
                // Folder일때 Level을 가져옴.
                if (Param01 != "S")
                {
                    //oForm.Items.Item("Level").Specific.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("Level").Specific.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                //공통 S
                oForm.Items.Item("Modual").Specific.Select(oRecordSet.Fields.Item("Modual").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub1").Specific.Select(oRecordSet.Fields.Item("Sub1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub2").Specific.Select(oRecordSet.Fields.Item("Sub2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub3").Specific.Select(oRecordSet.Fields.Item("Sub3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("No").Specific.Select(oRecordSet.Fields.Item("No").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Position").Specific.Select(oRecordSet.Fields.Item("Position").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("FatherID").Specific.Select(oRecordSet.Fields.Item("FatherID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("Strings").Value = oRecordSet.Fields.Item("Strings").Value;
                oForm.DataSources.UserDataSources.Item("UniqueId").Value = oRecordSet.Fields.Item("UniqueID").Value;
                oForm.Items.Item("Sequence").Specific.VALUE = oRecordSet.Fields.Item("Sequence").Value;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY998_SAVE
        /// </summary>
        private void PH_PY998_SAVE(int oRow = 0)
        {
            // 데이타 저장
            int ErrNum = 0;
            string sQry = string.Empty;
            string pGubun = string.Empty;
            string pFSGubun = string.Empty;
            string pUserID = string.Empty;
            string Modual = string.Empty;
            string Sub1 = string.Empty;
            string Sub2 = string.Empty;
            string Sub3 = string.Empty;
            string UserID = string.Empty;
            string No = string.Empty;
            string Level = string.Empty;
            string FatherID = string.Empty;
            string Position = string.Empty;
            string Strings_Renamed = string.Empty;
            string UniqueID = string.Empty;
            string pUniqueID = string.Empty;
            string Sequence = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry = "select UniqueID, Seq  from Authority_Screen where UserID ='";
                sQry = sQry + oForm.Items.Item("UserID").Specific.VALUE + "'";
                oRecordSet.DoQuery(sQry);

                for (int i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (oRecordSet.Fields.Item(0).Value == oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim())
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    if (oRecordSet.Fields.Item(1).Value == oForm.Items.Item("Sequence").Specific.VALUE.ToString().Trim())
                    {
                        ErrNum = 1;
                        throw new Exception();
                    }
                    oRecordSet.MoveNext();
                }

                if (PSH_Globals.SBO_Application.MessageBox("데이터 입력하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                pGubun = oForm.Items.Item("pGubun").Specific.VALUE.ToString().Trim();
                pFSGubun = oForm.Items.Item("pFSGubun").Specific.VALUE.ToString().Trim();
                pUserID = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();
                Modual = oForm.Items.Item("Modual").Specific.VALUE.ToString().Trim();
                Sub1 = oForm.Items.Item("Sub1").Specific.VALUE.ToString().Trim();
                Sub2 = oForm.Items.Item("Sub2").Specific.VALUE.ToString().Trim();
                Sub3 = oForm.Items.Item("Sub3").Specific.VALUE.ToString().Trim();
                UserID = oForm.Items.Item("UserID").Specific.VALUE.ToString().Trim();
                No = oForm.Items.Item("No").Specific.VALUE.ToString().Trim();
                Level = oForm.Items.Item("Level").Specific.VALUE.ToString().Trim();
                FatherID = oForm.Items.Item("FatherID").Specific.VALUE.ToString().Trim();
                Position = oForm.Items.Item("Position").Specific.VALUE.ToString().Trim();
                Strings_Renamed = oForm.Items.Item("Strings").Specific.VALUE.ToString().Trim();
                UniqueID = oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim();
                Sequence = oForm.Items.Item("Sequence").Specific.VALUE.ToString().Trim();
                pUniqueID = oForm.Items.Item("UniqueID").Specific.VALUE.ToString().Trim();

                sQry = "EXEC PH_PY998_03 '" + pFSGubun + "', '" + pUserID + "', '" + pUniqueID + "', '";
                sQry = sQry + Sequence + "', '" + UniqueID + "', '" + UserID + "', '" + FatherID + "', '" + Strings_Renamed + "', '";
                sQry = sQry + Position + "', '" + Level + "', '" + No + "', '" + pGubun + "', '";
                sQry = sQry + PSH_Globals.oCompany.UserName + "'";
                oDS_PH_PY998A.ExecuteQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("입력완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                PH_PY998_MTX01();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 저장된 값이 있습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력되지 않았습니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY998_Delete
        /// </summary>
        private void PH_PY998_Delete()
        {
            // 데이타 삭제
            string sQry = string.Empty;
            int ErrNum = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (PSH_Globals.SBO_Application.MessageBox("삭제하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if ((oForm.Items.Item("pFSGubun").Specific.VALUE == "F"))
                {
                    sQry = "delete from Authority_Folder where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                }
                else
                {
                    sQry = "delete from Authority_Screen where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                }
                oRecordSet.DoQuery(sQry);
                PH_PY998_MTX01();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("취소되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY998_Copy
        /// </summary>
        private void PH_PY998_Copy()
        {
            // 데이타 삭제
            string sQry = string.Empty;
            string pUserID = string.Empty;
            string CPUserID = string.Empty;
            int ErrNo = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            pUserID = oForm.Items.Item("pUserID").Specific.VALUE.ToString().Trim();
            CPUserID = oForm.Items.Item("CPUserID").Specific.VALUE.ToString().Trim();
            try
            {
                oForm.Freeze(true);

                sQry = "select count(1) from Authority_Screen where UserID ='" + CPUserID + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount <= 0)
                {
                    ErrNo = 1;
                    throw new Exception();
                }
                if (pUserID == "" || CPUserID == "")
                {
                    ErrNo = 2;
                    throw new Exception();
                }
                else
                {
                    sQry = "delete from Authority_Screen where UniqueID ='" + oForm.Items.Item("UniqueID").Specific.VALUE + "'";
                }

                sQry = "Insert into Authority_Screen";
                sQry = sQry + " select '" + CPUserID + "'";
                sQry = sQry + ", FatherID";
                sQry = sQry + ", String";
                sQry = sQry + ", UniqueID";
                sQry = sQry + ", Position";
                sQry = sQry + ", Type";
                sQry = sQry + ", Seq";
                sQry = sQry + ", Gubun";
                sQry = sQry + ", 'Y'";
                sQry = sQry + ", GETDATE()";
                sQry = sQry + ", '" + PSH_Globals.oCompany.UserName + "'";
                sQry = sQry + "  from Authority_Screen";
                sQry = sQry + "  where UserID ='" + pUserID + "'";

                oRecordSet.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                if (ErrNo == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("복제 계정에 권한이 있습니다. 복제 하고자 모든 권한을 삭제하세요.");
                }
                if (ErrNo == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("대상 ID와 복제ID는 필수입니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY998_Copy_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }
    }
}
