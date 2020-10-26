using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 재공 원가 이동등록
    /// </summary>
    internal class PS_CO160 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_CO160H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO160L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string oDocEntry;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO160.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO160_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO160");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_CO160_CreateItems();
                PS_CO160_ComboBox_Setting();
                PS_CO160_EnableMenus();
                PS_CO160_SetDocument(oFormDocEntry01);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_CO160_CreateItems()
        {
            try
            {
                //oForm.Freeze(true);

                oDS_PS_CO160H = oForm.DataSources.DBDataSources.Item("@PS_CO160H");
                oDS_PS_CO160L = oForm.DataSources.DBDataSources.Item("@PS_CO160L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_CO160_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //oForm.Freeze(true);

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //oForm.Freeze(false);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_CO160_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, true, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry01">DocEntry</param>
        private void PS_CO160_SetDocument(string oFormDocEntry01)
        {
            if (string.IsNullOrEmpty(oFormDocEntry01))
            {
                PS_CO160_FormItemEnabled();
                PS_CO160_AddMatrixRow(0, true);
            }
            else
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                PS_CO160_FormItemEnabled();
                oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO160_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("MItemCod").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oMat01.AutoResizeColumns();
                    PS_CO160_FormClear();
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("MItemCod").Enabled = true;
                    oForm.Items.Item("Comment").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;

                    oMat01.AutoResizeColumns();

                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("MItemCod").Enabled = false;

                    oMat01.AutoResizeColumns();

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_CO160_FormClear()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO160'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO160_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_CO160L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_CO160L.Offset = oRow;
                oDS_PS_CO160L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_CO160_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errCode = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value))
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                {
                    errCode = "2";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("MItemCod").Specific.Value))
                {
                    errCode = "3";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errCode = "4";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(i).Specific.Value))
                    {
                        errCode = "5";
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(i).Specific.Value))
                    {
                        errCode = "6";
                        throw new Exception();
                    }
                }

                oDS_PS_CO160L.RemoveRecord(oDS_PS_CO160L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_CO160_FormClear();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년월은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품목코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이동품목코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("MItemCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작지문서라인은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("PO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "6")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이동작지문서라인은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("MPO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_CO160_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_CO160_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                                dataHelpClass.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'I'");

                                PS_CO160_FormItemEnabled();
                                PS_CO160_AddMatrixRow(oMat01.RowCount, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                dataHelpClass.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'U'");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_CO160_FormItemEnabled();
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "MItemCod")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("MItemCod").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PO")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ColUID == "MPO")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                }
                else if (pVal.Before_Action == false)
                {
                }

                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            int i = 0;
            string errCode = string.Empty;
            string Query01 = string.Empty;
            //string ItemCode01 = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PO")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    errCode = "1";
                                    throw new Exception();
                                }
                                for (i = 1; i <= oMat01.RowCount; i++)
                                {
                                    if (pVal.Row != i) //현재 선택되어있는 행이 아니면
                                    {
                                        if (oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("PO").Cells.Item(i).Specific.Value)
                                        {
                                            errCode = "2";
                                            throw new Exception();
                                        }
                                    }
                                }

                                Query01 = "EXEC PS_CO160_01 '";
                                Query01 += oForm.Items.Item("BPLId").Specific.Value + "','";
                                Query01 += oForm.Items.Item("YM").Specific.Value + "','";
                                Query01 += oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value + "'";

                                oRecordSet.DoQuery(Query01);

                                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                                {
                                    oDS_PS_CO160L.SetValue("U_PO", pVal.Row - 1, oRecordSet.Fields.Item("PO").Value);
                                    oDS_PS_CO160L.SetValue("U_POEntry", pVal.Row - 1, oRecordSet.Fields.Item("POEntry").Value);
                                    oDS_PS_CO160L.SetValue("U_POLine", pVal.Row - 1, oRecordSet.Fields.Item("POLine").Value);
                                    oDS_PS_CO160L.SetValue("U_Sequence", pVal.Row - 1, oRecordSet.Fields.Item("Sequence").Value);
                                    oDS_PS_CO160L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item("ItemCode").Value);
                                    oDS_PS_CO160L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value);
                                    oDS_PS_CO160L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet.Fields.Item("CpCode").Value);
                                    oDS_PS_CO160L.SetValue("U_CpName", pVal.Row - 1, oRecordSet.Fields.Item("CpName").Value);
                                    oDS_PS_CO160L.SetValue("U_StcQty", pVal.Row - 1, oRecordSet.Fields.Item("StcQty").Value);
                                    oDS_PS_CO160L.SetValue("U_StcAmt", pVal.Row - 1, oRecordSet.Fields.Item("StcAmt").Value);
                                    oRecordSet.MoveNext();
                                }

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_CO160L.GetValue("U_PO", pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_CO160_AddMatrixRow(pVal.Row, false);
                                }

                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();
                            }
                            else if (pVal.ColUID == "MPO")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    errCode = "3";
                                    throw new Exception();
                                }
                                for (i = 1; i <= oMat01.RowCount; i++) //현재 선택되어있는 행이 아니면
                                {
                                    if (pVal.Row != i)
                                    {
                                        if (oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("MPO").Cells.Item(i).Specific.Value)
                                        {
                                            errCode = "4";
                                            throw new Exception();
                                        }
                                    }
                                }

                                Query01 = "EXEC PS_CO160_02 '";
                                Query01 += oForm.Items.Item("BPLId").Specific.Value + "','";
                                Query01 += oForm.Items.Item("YM").Specific.Value + "','";
                                Query01 += oForm.Items.Item("MItemCod").Specific.Value + "','";
                                Query01 += oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value + "'";

                                oRecordSet.DoQuery(Query01);

                                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                                {
                                    oDS_PS_CO160L.SetValue("U_MPO", pVal.Row - 1, oRecordSet.Fields.Item("MPO").Value);
                                    oDS_PS_CO160L.SetValue("U_MPOEntry", pVal.Row - 1, oRecordSet.Fields.Item("MPOEntry").Value);
                                    oDS_PS_CO160L.SetValue("U_MPOLine", pVal.Row - 1, oRecordSet.Fields.Item("MPOLine").Value);
                                    oDS_PS_CO160L.SetValue("U_MSequenc", pVal.Row - 1, oRecordSet.Fields.Item("MSequenc").Value);
                                    oDS_PS_CO160L.SetValue("U_MItemCod", pVal.Row - 1, oRecordSet.Fields.Item("MItemCod").Value);
                                    oDS_PS_CO160L.SetValue("U_MItemNam", pVal.Row - 1, oRecordSet.Fields.Item("MItemNam").Value);
                                    oDS_PS_CO160L.SetValue("U_MCpCode", pVal.Row - 1, oRecordSet.Fields.Item("MCpCode").Value);
                                    oDS_PS_CO160L.SetValue("U_MCpName", pVal.Row - 1, oRecordSet.Fields.Item("MCpName").Value);
                                    oRecordSet.MoveNext();
                                }

                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oDS_PS_CO160L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                    oDS_PS_CO160L.SetValue("U_Weight", pVal.Row - 1, "0");
                                    oDS_PS_CO160L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                }
                                else
                                {
                                    string ItemCode01 = oDS_PS_CO160L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();

                                    if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101") //EA자체품
                                    {
                                        oDS_PS_CO160L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102") //EAUOM
                                    {
                                        oDS_PS_CO160L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01))));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
                                    {
                                        oDS_PS_CO160L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
                                    {
                                        oDS_PS_CO160L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)));

                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG선택
                                    {
                                    }

                                    oDS_PS_CO160L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_CO160L.GetValue("U_Weight", pVal.Row - 1).ToString().Trim()) * Convert.ToDouble(oDS_PS_CO160L.GetValue("U_Price", pVal.Row - 1).ToString().Trim())));
                                    oDS_PS_CO160L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                            else
                            {
                                oDS_PS_CO160L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();

                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_CO160H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                oDS_PS_CO160H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_CO160H.SetValue("U_ItemName", 0, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ItemUID == "MItemCod")
                            {
                                oDS_PS_CO160H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_CO160H.SetValue("U_MItemNam", 0, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_CO160H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {

                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("동일한 항목이 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value = "";
                }
                else if (errCode == "3")
                {

                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("동일한 항목이 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value = "";
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO160_FormItemEnabled();
                    PS_CO160_AddMatrixRow(oMat01.VisualRowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO160H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO160L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "ItemName")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_CO160H", "U_ItemCode,U_ItemName", "", 0, "", "", "");
                    }

                    if (pVal.ItemUID == "MItemCod" || pVal.ItemUID == "MItemNam")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_CO160H", "U_MItemCod,U_MItemNam", "", 0, "", "", "");
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ROW_DELETE 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }

                        oMat01.FlushToDataSource();
                        oDS_PS_CO160L.RemoveRecord(oDS_PS_CO160L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_CO160_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_CO160L.GetValue("U_PO", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_CO160_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }

                        oForm.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 취소할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }

                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            dataHelpClass.DoQuery("EXEC PS_CO160_03 '" + oDocEntry + "', 'D'"); //삭제
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_CO160_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_CO160_FormItemEnabled();
                            PS_CO160_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
