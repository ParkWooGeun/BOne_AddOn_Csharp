using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 팀별 성과급차등 등급등록
    /// </summary>
    internal class PH_PY130 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY130A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY130B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        string g_preBankSel;

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY130.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY130_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY130");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm.DataBrowser.BrowseBy = "Code";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY130_CreateItems();
                PH_PY130_EnableMenus();
                PH_PY130_SetDocument(oFromDocEntry01);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY130_CreateItems()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oDS_PH_PY130A = oForm.DataSources.DBDataSources.Item("@PH_PY130A");
                oDS_PH_PY130B = oForm.DataSources.DBDataSources.Item("@PH_PY130B");

                oMat1 = oForm.Items.Item("Mat01").Specific;
                ////@PH_PY130B

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                ////----------------------------------------------------------------------------------------------
                //// 기본사항
                ////----------------------------------------------------------------------------------------------
                oForm.AutoManaged = true;
                dataHelpClass.AutoManaged(oForm,  "YM,JOBGBN,Number");

                //// 년월
                oDS_PH_PY130A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));

                //// 지급구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE CODE = 'P212' AND U_UseYN = 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific,"");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                ////사업장
                sQry = "SELECT BPLId, BPLName FROM [OBPL]";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!(oRecordSet.EoF))
                    {
                        oMat1.Columns.Item("CLTCOD").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("CLTCOD").DisplayDesc = true;


                ////등급
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '82' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!(oRecordSet.EoF))
                    {
                        oMat1.Columns.Item("Grade").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("Grade").DisplayDesc = true;

                oForm.DataSources.UserDataSources.Add("S_A", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_A").Specific.DataBind.SetBound(true, "", "S_A");

                oForm.DataSources.UserDataSources.Add("S_B", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_B").Specific.DataBind.SetBound(true, "", "S_B");

                oForm.DataSources.UserDataSources.Add("S_C", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_C").Specific.DataBind.SetBound(true, "", "S_C");

                oForm.DataSources.UserDataSources.Add("S_X", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_X").Specific.DataBind.SetBound(true, "", "S_X");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY130_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY130_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                 ////제거
                oForm.EnableMenu("1284", false);                ////취소
                oForm.EnableMenu("1293", true);                 ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY130_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY130_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY130_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY130_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY130_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY130_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.Items.Item("Btn01").Visible = true;

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", false);                    ////문서추가
                    oForm.EnableMenu("1293", false);                    ////행삭제
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("Btn01").Visible = true;


                    oForm.EnableMenu("1281", false);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가
                    oForm.EnableMenu("1293", false);                    ////행삭제
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("Btn01").Visible = false;

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가
                    oForm.EnableMenu("1293", false);                    ////행삭제

                }
                oForm.Items.Item("S_A").Specific.VALUE = 0;
                oForm.Items.Item("S_B").Specific.VALUE = 0;
                oForm.Items.Item("S_C").Specific.VALUE = 0;
                oForm.Items.Item("S_X").Specific.VALUE = 0;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY130_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
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

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                    //    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //        Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //        break;

                    //    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //        Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //        break;

                    //    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //        Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //        break;

                    //    case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //        Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //        break;

                    //    case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //        Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //        break;
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY130_DataValidCheck("Y") == false)
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
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PH_PY130_DataValidCheck("N") == false)
                        {
                            BubbleEvent = false;
                        }
                        else
                        {
                            if (PH_PY130_Create_Data() == false)
                            {
                                BubbleEvent = false;
                            }
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
                                PH_PY130_FormItemEnabled();
                                //                        Call PH_PY130_AddMatrixRow
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY130_FormItemEnabled();
                                //                        Call PH_PY130_AddMatrixRow
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                //                        Call PH_PY130_FormItemEnabled
                            }
                        }
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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

                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            int i = 0;
            int S_B = 0;
            int S_A = 0;
            int S_C = 0;
            int S_X = 0;
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
                        S_A = 0;
                        S_B = 0;
                        S_C = 0;
                        S_X = 0;

                        if (oMat1.VisualRowCount > 0)
                        {
                            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "A")
                                {
                                    S_A = S_A + 1;
                                }
                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "B")
                                {
                                    S_B = S_B + 1;
                                }
                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "C")
                                {
                                    S_C = S_C + 1;
                                }
                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "X")
                                {
                                    S_X = S_X + 1;
                                }

                            }
                            oForm.Items.Item("S_A").Specific.VALUE = S_A;
                            oForm.Items.Item("S_B").Specific.VALUE = S_B;
                            oForm.Items.Item("S_C").Specific.VALUE = S_C;
                            oForm.Items.Item("S_X").Specific.VALUE = S_X;
                        }
                        else
                        {
                            oForm.Items.Item("S_A").Specific.VALUE = 0;
                            oForm.Items.Item("S_B").Specific.VALUE = 0;
                            oForm.Items.Item("S_C").Specific.VALUE = 0;
                            oForm.Items.Item("S_X").Specific.VALUE = 0;
                        }
                    }
                }
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
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;

            int S_B = 0;
            int S_A = 0;
            int S_C = 0;
            int S_X = 0;
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat1.LoadFromDataSource();
                    PH_PY130_FormItemEnabled();


                    S_A = 0;
                    S_B = 0;
                    S_C = 0;
                    S_X = 0;

                    if (oMat1.VisualRowCount > 0)
                    {
                        for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                        {
                            if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "A")
                            {
                                S_A = S_A + 1;
                            }
                            else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "B")
                            {
                                S_B = S_B + 1;
                            }
                            else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "C")
                            {
                                S_C = S_C + 1;
                            }
                            else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "X")
                            {
                                S_X = S_X + 1;
                            }

                        }

                        oForm.Items.Item("S_A").Specific.VALUE = S_A;
                        oForm.Items.Item("S_B").Specific.VALUE = S_B;
                        oForm.Items.Item("S_C").Specific.VALUE = S_C;
                        oForm.Items.Item("S_X").Specific.VALUE = S_X;

                    }
                    else
                    {
                        oForm.Items.Item("S_A").Specific.VALUE = 0;
                        oForm.Items.Item("S_B").Specific.VALUE = 0;
                        oForm.Items.Item("S_C").Specific.VALUE = 0;
                        oForm.Items.Item("S_X").Specific.VALUE = 0;
                    }

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }


        /// <summary>
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY130A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY130B);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);
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
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY130_FormItemEnabled();
                            break;
                        //                Call PH_PY130_AddMatrixRow
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            ////문서찾기
                            PH_PY130_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY130_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY130_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            string sQry = string.Empty;

            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
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
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY130_DataValidCheck(string ChkYN)
        {
            bool functionReturnValue = false;

            functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            string tCode = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (ChkYN == "Y")
                {
                    if (string.IsNullOrEmpty(oDS_PH_PY130A.GetValue("U_JOBGBN", 0).ToString().Trim()))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }

                if (string.IsNullOrEmpty(oDS_PH_PY130A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY130A.GetValue("U_Number", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("회차는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Number").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //// 코드,이름 저장
                tCode = oDS_PH_PY130A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY130A.GetValue("U_JOBGBN", 0).ToString().Trim() + oDS_PH_PY130A.GetValue("U_Number", 0).ToString().Trim();
                oDS_PH_PY130A.SetValue("Code", 0, tCode);
                oDS_PH_PY130A.SetValue("Name", 0, tCode);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //// 데이터 중복 체크
                    sQry = "SELECT Code FROM [@PH_PY130A] WHERE Code = '" + tCode + "'";
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
                        PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 없습니다. 자료생성을 하기바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }

                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        ////구분
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Grade").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("등급은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("Grade").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY130_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY130_Create_Data
        /// </summary>
        public bool PH_PY130_Create_Data()
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                sQry = "exec [PH_PY130]";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if ((oRecordSet.RecordCount == 0))
                {
                    dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "E");
                    throw new Exception();
                }

                SAPbouiCOM.ProgressBar ProgressBar01 = null;
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY130B.InsertRecord((i));
                    }
                    oDS_PH_PY130B.Offset = i;
                    oDS_PH_PY130B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY130B.SetValue("U_CLTCOD", i, oRecordSet.Fields.Item("CLTCOD").Value);
                    oDS_PH_PY130B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item("TeamCode").Value);
                    oDS_PH_PY130B.SetValue("U_TeamName", i, oRecordSet.Fields.Item("TeamName").Value);
                    oRecordSet.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();
                ProgressBar01.Stop();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY130_Create_Data_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY130_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY130A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
                    functionReturnValue = false;
                    throw new Exception();
                }
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
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }
    }
}


//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
//// ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//    internal class PH_PY130
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY130.cls
//        ////  Module         : 급여관리 > 급여마스터관리
//        ////  Desc           : 팀별 성과급차등 등급등록
//        ////********************************************************************************

//        public string oFormUniqueID;
//        public SAPbouiCOM.Form oForm;

//        public SAPbouiCOM.Matrix oMat1;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY130A;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY130B;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public void LoadForm(string oFromDocEntry01 = "")
//        {

//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            // ERROR: Not supported in C#: OnErrorStatement


//            oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY130.srf");
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//            for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//            {
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//            }
//            oFormUniqueID = "PH_PY130_" + GetTotalFormsCount();
//            SubMain.AddForms(this, oFormUniqueID, "PH_PY130");
//            MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//            oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//            oForm.SupportedModes = -1;
//            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//            oForm.DataBrowser.BrowseBy = "Code";

//            oForm.Freeze(true);
//            PH_PY130_CreateItems();
//            PH_PY130_EnableMenus();
//            PH_PY130_SetDocument(oFromDocEntry01);
//            //    Call PH_PY130_FormResize

//            oForm.Update();
//            oForm.Freeze(false);

//            oForm.Visible = true;
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            return;
//        LoadForm_Error:

//            oForm.Update();
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oForm = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private bool PH_PY130_CreateItems()
//        {
//            bool functionReturnValue = false;

//            string sQry = null;
//            int i = 0;
//            string CLTCOD = null;

//            SAPbouiCOM.CheckBox oCheck = null;
//            SAPbouiCOM.EditText oEdit = null;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbouiCOM.Column oColumn = null;
//            SAPbouiCOM.Columns oColumns = null;
//            SAPbouiCOM.OptionBtn optBtn = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oDS_PH_PY130A = oForm.DataSources.DBDataSources("@PH_PY130A");
//            oDS_PH_PY130B = oForm.DataSources.DBDataSources("@PH_PY130B");


//            oMat1 = oForm.Items.Item("Mat01").Specific;
//            ////@PH_PY130B


//            oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            oMat1.AutoResizeColumns();


//            ////----------------------------------------------------------------------------------------------
//            //// 기본사항
//            ////----------------------------------------------------------------------------------------------
//            oForm.AutoManaged = true;
//            MDC_SetMod.AutoManaged(ref oForm, ref "YM,JOBGBN,Number");

//            //// 년월
//            oDS_PH_PY130A.SetValue("U_YM", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM"));

//            //// 지급구분
//            oCombo = oForm.Items.Item("JOBGBN").Specific;
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE CODE = 'P212' AND U_UseYN = 'Y'";
//            MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//            oForm.Items.Item("JOBGBN").DisplayDesc = true;

//            ////사업장
//            oColumn = oMat1.Columns.Item("CLTCOD");
//            sQry = "SELECT BPLId, BPLName FROM [OBPL]";
//            oRecordSet.DoQuery(sQry);
//            if (oRecordSet.RecordCount > 0)
//            {
//                while (!(oRecordSet.EoF))
//                {
//                    oColumn.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//                    oRecordSet.MoveNext();
//                }
//            }
//            oColumn.DisplayDesc = true;


//            ////등급
//            oColumn = oMat1.Columns.Item("Grade");
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '82' AND U_UseYN= 'Y'";
//            oRecordSet.DoQuery(sQry);
//            if (oRecordSet.RecordCount > 0)
//            {
//                while (!(oRecordSet.EoF))
//                {
//                    oColumn.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//                    oRecordSet.MoveNext();
//                }
//            }
//            oColumn.DisplayDesc = true;

//            oForm.DataSources.UserDataSources.Add("S_A", SAPbouiCOM.BoDataType.dt_SUM);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_A").Specific.DataBind.SetBound(true, "", "S_A");

//            oForm.DataSources.UserDataSources.Add("S_B", SAPbouiCOM.BoDataType.dt_SUM);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_B").Specific.DataBind.SetBound(true, "", "S_B");

//            oForm.DataSources.UserDataSources.Add("S_C", SAPbouiCOM.BoDataType.dt_SUM);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_C").Specific.DataBind.SetBound(true, "", "S_C");

//            oForm.DataSources.UserDataSources.Add("S_X", SAPbouiCOM.BoDataType.dt_SUM);
//            //UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_X").Specific.DataBind.SetBound(true, "", "S_X");


//            //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCheck = null;
//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            optBtn = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY130_CreateItems_Error:

//            //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCheck = null;
//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            optBtn = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private void PH_PY130_EnableMenus()
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.EnableMenu("1283", true);
//            ////제거
//            oForm.EnableMenu("1284", false);
//            ////취소
//            oForm.EnableMenu("1293", true);
//            ////행삭제

//            return;
//        PH_PY130_EnableMenus_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY130_SetDocument(string oFromDocEntry01)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((string.IsNullOrEmpty(oFromDocEntry01)))
//            {
//                PH_PY130_FormItemEnabled();
//                //        Call PH_PY130_AddMatrixRow
//            }
//            else
//            {
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                PH_PY130_FormItemEnabled();
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            return;
//        PH_PY130_SetDocument_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY130_FormItemEnabled()
//        {
//            SAPbouiCOM.ComboBox oCombo = null;

//            // ERROR: Not supported in C#: OnErrorStatement



//            oForm.Freeze(true);



//            if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//            {
//                oForm.Items.Item("Btn01").Visible = true;

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", false);
//                ////문서추가
//                oForm.EnableMenu("1293", false);
//                ////행삭제
//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//            {
//                oForm.Items.Item("Btn01").Visible = true;


//                oForm.EnableMenu("1281", false);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가
//                oForm.EnableMenu("1293", false);
//                ////행삭제
//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//            {
//                oForm.Items.Item("Btn01").Visible = false;

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가
//                oForm.EnableMenu("1293", false);
//                ////행삭제

//            }
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_A").Specific.VALUE = 0;
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_B").Specific.VALUE = 0;
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_C").Specific.VALUE = 0;
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("S_X").Specific.VALUE = 0;
//            oForm.Freeze(false);
//            return;
//        PH_PY130_FormItemEnabled_Error:

//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string sQry = null;
//            int i = 0;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            object S_B = null;
//            object S_A = null;
//            object S_C = null;
//            int S_X = 0;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                    ////1

//                    if (pVal.BeforeAction == true)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                            {
//                                if (PH_PY130_DataValidCheck(ref "Y") == false)
//                                {
//                                    BubbleEvent = false;
//                                }
//                                ////해야할일 작업
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                            {
//                                ////해야할일 작업

//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                            {

//                            }
//                        }
//                        if (pVal.ItemUID == "Btn01")
//                        {
//                            if (PH_PY130_DataValidCheck(ref "N") == false)
//                            {
//                                BubbleEvent = false;
//                            }
//                            else
//                            {
//                                if (PH_PY130_Create_Data() == false)
//                                {
//                                    BubbleEvent = false;
//                                }
//                            }
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    PH_PY130_FormItemEnabled();
//                                    //                        Call PH_PY130_AddMatrixRow
//                                }
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    PH_PY130_FormItemEnabled();
//                                    //                        Call PH_PY130_AddMatrixRow
//                                }
//                            }
//                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    //                        Call PH_PY130_FormItemEnabled
//                                }
//                            }
//                        }

//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    switch (pVal.ItemUID)
//                    {
//                        case "Mat01":
//                            if (pVal.Row > 0)
//                            {
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = pVal.ColUID;
//                                oLastColRow = pVal.Row;
//                            }
//                            break;
//                        default:
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = "";
//                            oLastColRow = 0;
//                            break;
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//                    ////4
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                    ////5
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {


//                            //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            S_A = 0;
//                            //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            S_B = 0;
//                            //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            S_C = 0;
//                            S_X = 0;

//                            if (oMat1.VisualRowCount > 0)
//                            {
//                                for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
//                                {
//                                    //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "A")
//                                    {
//                                        //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        S_A = S_A + 1;
//                                        //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    }
//                                    else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "B")
//                                    {
//                                        //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        S_B = S_B + 1;
//                                        //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    }
//                                    else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "C")
//                                    {
//                                        //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        S_C = S_C + 1;
//                                        //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    }
//                                    else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "X")
//                                    {
//                                        S_X = S_X + 1;
//                                    }

//                                }

//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_A").Specific.VALUE = S_A;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_B").Specific.VALUE = S_B;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_C").Specific.VALUE = S_C;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_X").Specific.VALUE = S_X;

//                            }
//                            else
//                            {
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_A").Specific.VALUE = 0;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_B").Specific.VALUE = 0;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_C").Specific.VALUE = 0;
//                                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                oForm.Items.Item("S_X").Specific.VALUE = 0;
//                            }



//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    if (pVal.BeforeAction == true)
//                    {
//                        switch (pVal.ItemUID)
//                        {
//                            case "Mat01":
//                                if (pVal.Row > 0)
//                                {
//                                    oMat1.SelectRow(pVal.Row, true, false);
//                                }
//                                break;
//                        }

//                        switch (pVal.ItemUID)
//                        {
//                            case "Mat01":
//                                if (pVal.Row > 0)
//                                {
//                                    oLastItemUID = pVal.ItemUID;
//                                    oLastColUID = pVal.ColUID;
//                                    oLastColRow = pVal.Row;
//                                }
//                                break;
//                            default:
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = "";
//                                oLastColRow = 0;
//                                break;
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//                    ////7
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//                    ////8
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//                    ////9
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                    ////10
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {

//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    if (pVal.BeforeAction == true)
//                    {
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        oMat1.LoadFromDataSource();
//                        PH_PY130_FormItemEnabled();


//                        //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        S_A = 0;
//                        //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        S_B = 0;
//                        //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        S_C = 0;
//                        S_X = 0;

//                        if (oMat1.VisualRowCount > 0)
//                        {
//                            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
//                            {
//                                //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "A")
//                                {
//                                    //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    S_A = S_A + 1;
//                                    //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                }
//                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "B")
//                                {
//                                    //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    S_B = S_B + 1;
//                                    //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                }
//                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "C")
//                                {
//                                    //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                    S_C = S_C + 1;
//                                    //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i + 1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                }
//                                else if (oMat1.Columns.Item("Grade").Cells.Item(i + 1).Specific.VALUE == "X")
//                                {
//                                    S_X = S_X + 1;
//                                }

//                            }

//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: S_A 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_A").Specific.VALUE = S_A;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: S_B 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_B").Specific.VALUE = S_B;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            //UPGRADE_WARNING: S_C 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_C").Specific.VALUE = S_C;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_X").Specific.VALUE = S_X;

//                        }
//                        else
//                        {
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_A").Specific.VALUE = 0;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_B").Specific.VALUE = 0;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_C").Specific.VALUE = 0;
//                            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item("S_X").Specific.VALUE = 0;
//                        }

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//                    ////12
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//                    ////16
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                    ////17
//                    if (pVal.BeforeAction == true)
//                    {
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        SubMain.RemoveForms(oFormUniqueID);
//                        //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oForm = null;
//                        //UPGRADE_NOTE: oDS_PH_PY130A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY130A = null;
//                        //UPGRADE_NOTE: oDS_PH_PY130B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY130B = null;

//                        //UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oMat1 = null;

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//                    ////18
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//                    ////19
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//                    ////20
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//                    ////21
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        //                oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10
//                        //                oForm.Items.Item("79").Height = oForm.Items.Item("80").Height
//                        //
//                        //                oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16
//                        //                oForm.Items.Item("77").Height = oForm.Items.Item("78").Height
//                        //
//                        //                oForm.Items.Item("8").Width = oForm.Items.Item("Mat02").Left + oForm.Items.Item("Mat02").Width + 5
//                        //                oForm.Items.Item("8").Height = oForm.Items.Item("1").Top - 80
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//                    ////22
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//                    ////23
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//                    ////27
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.Before_Action == false)
//                    {
//                        //                If pVal.ItemUID = "Code" Then
//                        //                    Call MDC_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY130A", "Code")
//                        //                End If
//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//                    ////37
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//                    ////38
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_Drag:
//                    ////39
//                    break;

//            }

//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;

//            return;
//        Raise_FormItemEvent_Error:
//            ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            oForm.Freeze((false));
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        {
//            int i = 0;
//            // ERROR: Not supported in C#: OnErrorStatement

//            oForm.Freeze(true);

//            if ((pVal.BeforeAction == true))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    case "1293":
//                        break;
//                    case "1281":
//                        break;
//                    case "1282":
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        break;
//                        //                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY130A", "Code")      '//접속자 권한에 따른 사업장 보기
//                }
//            }
//            else if ((pVal.BeforeAction == false))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        PH_PY130_FormItemEnabled();
//                        break;
//                    //                Call PH_PY130_AddMatrixRow
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    //            Case "1293":
//                    //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//                    case "1281":
//                        ////문서찾기
//                        PH_PY130_FormItemEnabled();
//                        //                Call PH_PY130_AddMatrixRow
//                        oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        break;
//                    case "1282":
//                        ////문서추가
//                        PH_PY130_FormItemEnabled();
//                        break;
//                    //                Call PH_PY130_AddMatrixRow
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        PH_PY130_FormItemEnabled();
//                        break;
//                    case "1293":
//                        //// 행삭제
//                        break;

//                        //                Call PH_PY130_AddMatrixRow
//                }
//            }
//            oForm.Freeze(false);
//            return;
//        Raise_FormMenuEvent_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((BusinessObjectInfo.BeforeAction == true))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            else if ((BusinessObjectInfo.BeforeAction == false))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;

//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            return;
//        Raise_FormDataEvent_Error:


//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//            }
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
//            return;
//        Raise_RightClickEvent_Error:

//            MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY130_AddMatrixRow()
//        {
//            int oRow = 0;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            ////[Mat1]
//            oMat1.FlushToDataSource();
//            oRow = oMat1.VisualRowCount;

//            if (oMat1.VisualRowCount > 0)
//            {
//                if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY130B.GetValue("U_CLTCOD", oRow - 1))))
//                {
//                    if (oDS_PH_PY130B.Size <= oMat1.VisualRowCount)
//                    {
//                        oDS_PH_PY130B.InsertRecord((oRow));
//                    }
//                    oDS_PH_PY130B.Offset = oRow;
//                    oDS_PH_PY130B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                    oDS_PH_PY130B.SetValue("U_Date", oRow, "");
//                    oDS_PH_PY130B.SetValue("U_Type", oRow, "");
//                    oDS_PH_PY130B.SetValue("U_Comments", oRow, Convert.ToString(0));
//                    oDS_PH_PY130B.SetValue("U_Close", oRow, Convert.ToString(0));
//                    oMat1.LoadFromDataSource();
//                }
//                else
//                {
//                    oDS_PH_PY130B.Offset = oRow - 1;
//                    oDS_PH_PY130B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//                    oDS_PH_PY130B.SetValue("U_Date", oRow - 1, "");
//                    oDS_PH_PY130B.SetValue("U_Type", oRow - 1, "");
//                    oDS_PH_PY130B.SetValue("U_Comments", oRow - 1, Convert.ToString(0));
//                    oDS_PH_PY130B.SetValue("U_Close", oRow - 1, Convert.ToString(0));
//                    oMat1.LoadFromDataSource();
//                }
//            }
//            else if (oMat1.VisualRowCount == 0)
//            {
//                oDS_PH_PY130B.Offset = oRow;
//                oDS_PH_PY130B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oDS_PH_PY130B.SetValue("U_Date", oRow, "");
//                oDS_PH_PY130B.SetValue("U_Type", oRow, "");
//                oDS_PH_PY130B.SetValue("U_Comments", oRow, Convert.ToString(0));
//                oDS_PH_PY130B.SetValue("U_Close", oRow, Convert.ToString(0));
//                oMat1.LoadFromDataSource();
//            }

//            oForm.Freeze(false);
//            return;
//        PH_PY130_AddMatrixRow_Error:
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY130_FormClear()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            string DocEntry = null;
//            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY130'", ref "");
//            if (Convert.ToDouble(DocEntry) == 0)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//            }
//            else
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//            }
//            return;
//        PH_PY130_FormClear_Error:
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public bool PH_PY130_DataValidCheck(ref string ChkYN)
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = false;
//            int i = 0;
//            string sQry = null;
//            string tCode = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            if (ChkYN == "Y")
//            {
//                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY130A.GetValue("U_JOBGBN", 0))))
//                {
//                    MDC_Globals.Sbo_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                }
//            }

//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY130A.GetValue("U_YM", 0))))
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY130A.GetValue("U_Number", 0))))
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("회차는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("Number").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                functionReturnValue = false;
//                return functionReturnValue;
//            }

//            //// 코드,이름 저장
//            tCode = Strings.Trim(oDS_PH_PY130A.GetValue("U_YM", 0)) + Strings.Trim(oDS_PH_PY130A.GetValue("U_JOBGBN", 0)) + Strings.Trim(oDS_PH_PY130A.GetValue("U_Number", 0));
//            oDS_PH_PY130A.SetValue("Code", 0, tCode);
//            oDS_PH_PY130A.SetValue("Name", 0, tCode);

//            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//            {
//                //// 데이터 중복 체크
//                sQry = "SELECT Code FROM [@PH_PY130A] WHERE Code = '" + tCode + "'";
//                oRecordSet.DoQuery(sQry);

//                if (oRecordSet.RecordCount > 0)
//                {
//                    MDC_Globals.Sbo_Application.SetStatusBarMessage("이미 데이터가 존재합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                }
//            }

//            if (ChkYN == "Y")
//            {
//                if (oMat1.VisualRowCount == 0)
//                {
//                    MDC_Globals.Sbo_Application.SetStatusBarMessage("데이터가 없습니다. 자료생성을 하기바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                }

//                for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
//                {
//                    ////구분
//                    //UPGRADE_WARNING: oMat1.Columns(Grade).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    if (string.IsNullOrEmpty(oMat1.Columns.Item("Grade").Cells.Item(i).Specific.VALUE))
//                    {
//                        MDC_Globals.Sbo_Application.SetStatusBarMessage("등급은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                        oMat1.Columns.Item("Grade").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        functionReturnValue = false;
//                        return functionReturnValue;
//                    }
//                }
//            }
//            else
//            {

//            }



//            functionReturnValue = true;
//            return functionReturnValue;


//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//        PH_PY130_DataValidCheck_Error:


//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            functionReturnValue = false;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }


//        public bool PH_PY130_Create_Data()
//        {
//            bool functionReturnValue = false;
//            int i = 0;
//            string sQry = null;


//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//            sQry = "exec [PH_PY130]";
//            oRecordSet.DoQuery(sQry);

//            oMat1.Clear();
//            oMat1.FlushToDataSource();
//            oMat1.LoadFromDataSource();

//            if ((oRecordSet.RecordCount == 0))
//            {
//                MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//                goto PH_PY130_Create_Data_Exit;
//            }

//            SAPbouiCOM.ProgressBar ProgressBar01 = null;
//            ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
//            {
//                if (i != 0)
//                {
//                    oDS_PH_PY130B.InsertRecord((i));
//                }
//                oDS_PH_PY130B.Offset = i;
//                oDS_PH_PY130B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                oDS_PH_PY130B.SetValue("U_CLTCOD", i, oRecordSet.Fields.Item("CLTCOD").Value);
//                oDS_PH_PY130B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item("TeamCode").Value);
//                oDS_PH_PY130B.SetValue("U_TeamName", i, oRecordSet.Fields.Item("TeamName").Value);
//                oRecordSet.MoveNext();
//                ProgressBar01.Value = ProgressBar01.Value + 1;
//                ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
//            }
//            oMat1.LoadFromDataSource();
//            oMat1.AutoResizeColumns();
//            oForm.Update();

//            ProgressBar01.Stop();
//            //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgressBar01 = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY130_Create_Data_Exit:
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            if ((ProgressBar01 != null))
//            {
//                ProgressBar01.Stop();
//            }
//        PH_PY130_Create_Data_Error:
//            ProgressBar01.Stop();
//            //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            ProgressBar01 = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_Create_Data_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        public bool PH_PY130_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = true;
//            object i = null;
//            int j = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            //UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY130A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY130A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
//            {
//                MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                goto PH_PY130_Validate_Exit;
//            }
//            //
//            if (ValidateType == "수정")
//            {

//            }
//            else if (ValidateType == "행삭제")
//            {

//            }
//            else if (ValidateType == "취소")
//            {

//            }
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY130_Validate_Exit:
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY130_Validate_Error:
//            functionReturnValue = false;
//            MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY130_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }
//    }
//}
