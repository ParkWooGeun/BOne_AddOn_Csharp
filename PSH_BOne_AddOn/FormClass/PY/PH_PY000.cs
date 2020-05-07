using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사원마스터등록(사용안됨)
    /// </summary>
    internal class PH_PY000 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY000A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY000B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        string g_preBankSel;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY000.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY000_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY000");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY000_CreateItems();
                PH_PY000_EnableMenus();
                PH_PY000_SetDocument(oFromDocEntry01);
                oForm.Update();
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
        private void PH_PY000_CreateItems()
        {
            int i = 0;
            string sQry = string.Empty;
            string sTemp = string.Empty;
            string sTable = string.Empty;
            string sField1 = string.Empty;
            string sField2 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY000A = oForm.DataSources.DBDataSources.Item("@PH_PY000A");
                oDS_PH_PY000B = oForm.DataSources.DBDataSources.Item("@PH_PY000B");

                oMat1 = oForm.Items.Item("Mat1").Specific;
                ////@PH_PY000B

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                

                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("Value").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("Value").DisplayDesc = true;
                oMat1.Columns.Item("Value").Editable = true;

                oForm.AutoManaged = true;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY000_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY000_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY000_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY000_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY000_FormItemEnabled();
                    PH_PY000_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY000_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY000_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        /// <summary>
        /// PH_PY000_FormItemEnabled 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY000_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", false);                   // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.EnableMenu("1281", false);                   // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY000_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY000_AddMatrixRow 매트릭스 행 추가
        /// </summary>
        public void PH_PY000_AddMatrixRow()
        {
            int oRow = 0;
            try
            {
                oForm.Freeze(true);
                ////[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY000B.GetValue("U_UserCode", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY000B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY000B.InsertRecord((oRow));
                        }
                        oDS_PH_PY000B.Offset = oRow;
                        oDS_PH_PY000B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY000B.SetValue("U_UserCode", oRow, "");
                        oDS_PH_PY000B.SetValue("U_UserName", oRow, "");
                        oDS_PH_PY000B.SetValue("U_Value", oRow, "");
                        oDS_PH_PY000B.SetValue("U_Default", oRow, "");
                        oDS_PH_PY000B.SetValue("U_Comments", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY000B.Offset = oRow - 1;
                        oDS_PH_PY000B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY000B.SetValue("U_UserCode", oRow - 1, "");
                        oDS_PH_PY000B.SetValue("U_UserName", oRow - 1, "");
                        oDS_PH_PY000B.SetValue("U_value", oRow - 1, "");
                        oDS_PH_PY000B.SetValue("U_Default", oRow - 1, "");
                        oDS_PH_PY000B.SetValue("U_Comments", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY000B.Offset = oRow;
                    oDS_PH_PY000B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY000B.SetValue("U_UserCode", oRow, "");
                    oDS_PH_PY000B.SetValue("U_UserName", oRow, "");
                    oDS_PH_PY000B.SetValue("U_Value", oRow, "");
                    oDS_PH_PY000B.SetValue("U_Default", oRow, "");
                    oDS_PH_PY000B.SetValue("U_Comments", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY000_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    case "Mat1":
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
        /// Raise_EVENT_ROW_DELETE 
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="oMat"></param>
        /// <param name="DBData"></param>
        /// <param name="CheckField"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;
            try
            {
                if ((oLastColRow > 0))
                {
                    if (pVal.BeforeAction == true)
                    {

                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while ((i <= DBData.Size - 1))
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord((i));
                                    i = 0;
                                }
                                else
                                {
                                    i = i + 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case "Mat1":
                        case "Grid1":
                            if (pVal.Row > 0)
                            {
                                ////Call oMat1.SelectRow(pval.Row, True, False)

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
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat1.LoadFromDataSource();

                    PH_PY000_FormItemEnabled();
                    PH_PY000_AddMatrixRow();

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
        /// Raise_EVENT_VALIDATE이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
            string sQry = string.Empty;
            string sTemp = string.Empty;
            string sTable = string.Empty;
            string sField1 = string.Empty;
            string sField2 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "UserCode")
                        {
                            if (!string.IsNullOrEmpty(oMat1.Columns.Item("UserCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                sTemp = dataHelpClass.Get_ReData("U_Name", "User_Code", "OUSR", "'" + oMat1.Columns.Item("UserCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'","");
                                if (!string.IsNullOrEmpty(sTemp))
                                {
                                    oMat1.Columns.Item("UserName").Cells.Item(pVal.Row).Specific.Value = sTemp;
                                    oMat1.AutoResizeColumns();
                                }
                                else
                                {
                                    oMat1.Columns.Item("UserName").Cells.Item(pVal.Row).Specific.Value = "";
                                    BubbleEvent = false;
                                    oMat1.AutoResizeColumns();
                                }
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "UserCode")
                        {
                            PH_PY000_AddMatrixRow();
                            oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        if (pVal.ItemUID == "CodField" | pVal.ItemUID == "NamField")
                        {
                            sTable = oForm.Items.Item("DefTable").Specific.Value;
                            sField1 = oForm.Items.Item("CodField").Specific.Value;
                            sField2 = oForm.Items.Item("NamField").Specific.Value;

                            if (!string.IsNullOrEmpty(sTable) & !string.IsNullOrEmpty(sField1) & !string.IsNullOrEmpty(sField2))
                            {
                                if (oMat1.Columns.Item("Value").ValidValues.Count > 0)
                                {
                                    for (i = oMat1.Columns.Item("Value").ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oMat1.Columns.Item("Value").ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }

                                if (dataHelpClass.TableFieldCheck(sTable, sField1, sField2) == true)
                                {
                                    sQry = "Select " + sField1 + "," + sField2 + " From " + sTable;
                                    oRecordSet.DoQuery(sQry);
                                    if (oRecordSet.RecordCount > 0)
                                    {
                                        while (!(oRecordSet.EoF))
                                        {
                                            oMat1.Columns.Item("Value").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                            oRecordSet.MoveNext();
                                        }
                                    }
                                }

                            }
                            else
                            {
                                if (oMat1.Columns.Item("Value").ValidValues.Count > 0)
                                {
                                    for (i = oMat1.Columns.Item("Value").ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oMat1.Columns.Item("Value").ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oMat1.Columns.Item("Value").ValidValues.Add("", "");
                                }
                            }
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY000A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY000B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
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
                            if (PH_PY000_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY000_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
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
                                //                            Call PH_PY000_FormItemEnabled
                                //                            Call PH_PY000_AddMatrixRow
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                //                            Call PH_PY000_FormItemEnabled
                                //                            Call PH_PY000_AddMatrixRow
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                //                            Call PH_PY000_FormItemEnabled
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
        /// Raise_FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            int i = 0;
            string sQry = string.Empty;
            string sTable = string.Empty;
            string sField1 = string.Empty;
            string sField2 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:  // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: // 36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: // 33
                            sTable = oForm.Items.Item("DefTable").Specific.Value.ToString().Trim();
                            sField1 = oForm.Items.Item("CodField").Specific.Value.ToString().Trim();
                            sField2 = oForm.Items.Item("NamField").Specific.Value.ToString().Trim();

                            if (!string.IsNullOrEmpty(sTable) & !string.IsNullOrEmpty(sField1) & !string.IsNullOrEmpty(sField2))
                            {
                                if (oMat1.Columns.Item("Value").ValidValues.Count > 0)
                                {
                                    for (i = oMat1.Columns.Item("Value").ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oMat1.Columns.Item("Value").ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }

                                if (dataHelpClass.TableFieldCheck(sTable, sField1, sField2) == true)
                                {
                                    sQry = "Select " + sField1 + "," + sField2 + " From " + sTable;
                                    oRecordSet.DoQuery(sQry);
                                    if (oRecordSet.RecordCount > 0)
                                    {
                                        while (!(oRecordSet.EoF))
                                        {
                                            oMat1.Columns.Item("Value").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                            oRecordSet.MoveNext();
                                        }
                                    }
                                }

                            }
                            else
                            {
                                if (oMat1.Columns.Item("Value").ValidValues.Count > 0)
                                {
                                    for (i = oMat1.Columns.Item("Value").ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oMat1.Columns.Item("Value").ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oMat1.Columns.Item("Value").ValidValues.Add("", "");
                                }
                            }
                            oMat1.AutoResizeColumns();
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            // 36
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
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
                            PH_PY000_FormItemEnabled();
                            PH_PY000_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY000_FormItemEnabled();
                            PH_PY000_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY000_FormItemEnabled();
                            PH_PY000_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY000_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제
                            // [MAT1 용]
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY000B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY000B.GetValue("U_UserCode", i)))
                                    {
                                        oDS_PH_PY000B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY000B.Size; i++)
                                {
                                    oDS_PH_PY000B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY000_AddMatrixRow();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY000_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oMat1.FlushToDataSource();
                // Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY000B.Size > 1)
                    oDS_PH_PY000B.RemoveRecord((oDS_PH_PY000B.Size - 1));
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY000_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            return functionReturnValue;
        }
    }
}




//using System;

//using SAPbouiCOM;
//using PSH_BOne_AddOn;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 사원마스터등록(사용안됨)
//    /// </summary>
//    internal class PH_PY000 : PSH_BaseClass
//    {
//        public string oFormUniqueID;
//        //public SAPbouiCOM.Form oForm;

//        public SAPbouiCOM.Matrix oMat1;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY000A;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY000B;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        /// <summary>
//        /// Form 호출
//        /// </summary>
//        /// <param name="oFromDocEntry01"></param>
//        public override void LoadForm(string oFromDocEntry01 = "")
//        {
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY000.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (int loopCount = 1; loopCount <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); loopCount++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[loopCount - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[loopCount - 1].nodeValue = 16;
//                }
//                oFormUniqueID = "PH_PY000_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY000");

//                string strXml = null;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                oForm.DataBrowser.BrowseBy = "Code";

//                oForm.Freeze(true);
//                //PH_PY000_CreateItems();
//                //PH_PY000_EnableMenus();
//                //PH_PY000_SetDocument(oFromDocEntry01);
//                //    Call PH_PY000_FormResize
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("LoadForm_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
//            }
//        }

//		private bool PH_PY000_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbouiCOM.OptionBtn optBtn = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			////----------------------------------------------------------------------------------------------
//			//// 데이터셋정의
//			////----------------------------------------------------------------------------------------------
//			////테이블이 있을경우 데이터셋(Matrix)
//			oDS_PH_PY000A = oForm.DataSources.DBDataSources("@PH_PY000A");
//			////헤더
//			oDS_PH_PY000B = oForm.DataSources.DBDataSources("@PH_PY000B");
//			////라인

//			oMat1 = oForm.Items.Item("Mat1").Specific;
//			//

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();


//			oColumn = oMat1.Columns.Item("Value");
//			oColumn.DisplayDesc = true;
//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------


//			////매트릭스컬럼
//			oColumn = oMat1.Columns.Item("Value");
//			oColumn.Editable = true;

//			//    '//UserDataSources
//			//    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
//			//    Set oCombo = oForm.Items("    ").Specific
//			//    oCombo.DataBind.SetBound True, "", "    "
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			//    oForm.Items("CLTCOD").DisplayDesc = True

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY000_CreateItems_Error:

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY000_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY000_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY000_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY000_FormItemEnabled();
//				PH_PY000_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY000_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY000_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY000_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {


//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY000_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			short ErrNum = 0;
//			string sTemp = null;

//			string sTable = null;
//			string sField1 = null;
//			string sField2 = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			ErrNum = 0;
//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY000_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (PH_PY000_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY000_FormItemEnabled
//									//                            Call PH_PY000_AddMatrixRow
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY000_FormItemEnabled
//									//                            Call PH_PY000_AddMatrixRow
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY000_FormItemEnabled
//								}
//							}
//						}

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat1":
//						case "Grid1":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					break;
//				//            Call oForm.Freeze(True)
//				//            If pval.BeforeAction = True Then
//				//
//				//            ElseIf pval.BeforeAction = False Then
//				//                If pval.ItemChanged = True Then
//				//
//				//                End If
//				//            End If
//				//            Call oForm.Freeze(False)
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat1":
//							case "Grid1":
//								if (pval.Row > 0) {
//									////Call oMat1.SelectRow(pval.Row, True, False)

//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "Mat1" & pval.ColUID == "UserCode") {

//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "Mat1" & pval.ColUID == "UserCode") {
//								//UPGRADE_WARNING: oMat1.Columns(UserCode).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (!string.IsNullOrEmpty(oMat1.Columns.Item("UserCode").Cells.Item(pval.Row).Specific.Value)) {
//									//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sTemp = MDC_SetMod.Get_ReData("U_Name", "User_Code", "OUSR", "'" + Strings.Trim(oMat1.Columns.Item("UserCode").Cells.Item(pval.Row).Specific.Value) + "'");
//									if (!string.IsNullOrEmpty(sTemp)) {
//										//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oMat1.Columns.Item("UserName").Cells.Item(pval.Row).Specific.Value = sTemp;
//										oMat1.AutoResizeColumns();
//									} else {
//										//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//										oMat1.Columns.Item("UserName").Cells.Item(pval.Row).Specific.Value = "";
//										BubbleEvent = false;
//										oMat1.AutoResizeColumns();
//									}
//								}
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "Mat1" & pval.ColUID == "UserCode") {
//								PH_PY000_AddMatrixRow();
//								oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							}
//							if (pval.ItemUID == "CodField" | pval.ItemUID == "NamField") {
//								oColumn = oMat1.Columns.Item("Value");
//								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sTable = oForm.Items.Item("DefTable").Specific.Value;
//								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sField1 = oForm.Items.Item("CodField").Specific.Value;
//								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sField2 = oForm.Items.Item("NamField").Specific.Value;

//								if (!string.IsNullOrEmpty(sTable) & !string.IsNullOrEmpty(sField1) & !string.IsNullOrEmpty(sField2)) {
//									if (oColumn.ValidValues.Count > 0) {
//										for (i = oColumn.ValidValues.Count - 1; i >= 0; i += -1) {
//											oColumn.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//										}
//									}

//									if (MDC_SetMod.TableFieldCheck(ref sTable, ref sField1, ref sField2) == true) {
//										sQry = "Select " + sField1 + "," + sField2 + " From " + sTable;
//										oRecordSet.DoQuery(sQry);
//										if (oRecordSet.RecordCount > 0) {
//											while (!(oRecordSet.EoF)) {
//												oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//												oRecordSet.MoveNext();
//											}
//										}
//									}

//								} else {
//									if (oColumn.ValidValues.Count > 0) {
//										for (i = oColumn.ValidValues.Count - 1; i >= 0; i += -1) {
//											oColumn.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//										}
//										oColumn.ValidValues.Add("", "");
//									}
//								}
//							}
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY000_FormItemEnabled();
//						PH_PY000_AddMatrixRow();
//						//
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY000A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY000A = null;
//						//UPGRADE_NOTE: oDS_PH_PY000B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY000B = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//						//                Set oGrid1 = Nothing

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					break;
//				//            If pval.BeforeAction = True Then
//				//
//				//            ElseIf pval.BeforeAction = False Then
//				//
//				//            End If
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					break;
//				//            If pval.BeforeAction = True Then
//				//
//				//            ElseIf pval.Before_Action = False Then
//				//
//				//            End If
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;

//			}
//			ErrMessage:

//			if (ErrNum == 1) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("테이블이 존재 하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
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
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY000_FormItemEnabled();
//						PH_PY000_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY000_FormItemEnabled();
//						PH_PY000_AddMatrixRow();
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY000_FormItemEnabled();
//						PH_PY000_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY000_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						//// [MAT1 용]
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY000B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY000B.GetValue("U_UserCode", i))) {
//									oDS_PH_PY000B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY000B.Size; i++) {
//								oDS_PH_PY000B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY000_AddMatrixRow();
//						break;
//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			SAPbouiCOM.Column oColumn = null;
//			string sTable = null;
//			string sField1 = null;
//			string sField2 = null;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == true)) {
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
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						oColumn = oMat1.Columns.Item("Value");
//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sTable = oForm.Items.Item("DefTable").Specific.Value;
//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sField1 = oForm.Items.Item("CodField").Specific.Value;
//						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sField2 = oForm.Items.Item("NamField").Specific.Value;

//						if (!string.IsNullOrEmpty(sTable) & !string.IsNullOrEmpty(sField1) & !string.IsNullOrEmpty(sField2)) {
//							if (oColumn.ValidValues.Count > 0) {
//								for (i = oColumn.ValidValues.Count - 1; i >= 0; i += -1) {
//									oColumn.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//							}

//							if (MDC_SetMod.TableFieldCheck(ref sTable, ref sField1, ref sField2) == true) {
//								sQry = "Select " + sField1 + "," + sField2 + " From " + sTable;
//								oRecordSet.DoQuery(sQry);
//								if (oRecordSet.RecordCount > 0) {
//									while (!(oRecordSet.EoF)) {
//										oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//										oRecordSet.MoveNext();
//									}
//								}
//							}

//						} else {
//							if (oColumn.ValidValues.Count > 0) {
//								for (i = oColumn.ValidValues.Count - 1; i >= 0; i += -1) {
//									oColumn.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//								}
//								oColumn.ValidValues.Add("", "");
//							}
//						}
//						oMat1.AutoResizeColumns();
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
//			return;
//			Raise_FormDataEvent_Error:


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat1":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY000_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1 용]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY000B.GetValue("U_UserCode", oRow - 1)))) {
//					if (oDS_PH_PY000B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY000B.InsertRecord((oRow));
//					}
//					oDS_PH_PY000B.Offset = oRow;
//					oDS_PH_PY000B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY000B.SetValue("U_UserCode", oRow, "");
//					oDS_PH_PY000B.SetValue("U_UserName", oRow, "");
//					oDS_PH_PY000B.SetValue("U_Value", oRow, "");
//					oDS_PH_PY000B.SetValue("U_Default", oRow, "");
//					oDS_PH_PY000B.SetValue("U_Comments", oRow, "");
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY000B.Offset = oRow - 1;
//					oDS_PH_PY000B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY000B.SetValue("U_UserCode", oRow - 1, "");
//					oDS_PH_PY000B.SetValue("U_UserName", oRow - 1, "");
//					oDS_PH_PY000B.SetValue("U_value", oRow - 1, "");
//					oDS_PH_PY000B.SetValue("U_value", oRow - 1, "");
//					oDS_PH_PY000B.SetValue("U_value", oRow - 1, "");
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY000B.Offset = oRow;
//				oDS_PH_PY000B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY000B.SetValue("U_UserCode", oRow, "");
//				oDS_PH_PY000B.SetValue("U_UserName", oRow, "");
//				oDS_PH_PY000B.SetValue("U_Value", oRow, "");
//				oDS_PH_PY000B.SetValue("U_Default", oRow, "");
//				oDS_PH_PY000B.SetValue("U_Comments", oRow, "");
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY000_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY000_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY000'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//			}
//			return;
//			PH_PY000_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY000_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			int iLooper = 0;
//			SAPbouiCOM.Matrix vMatrix = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////----------------------------------------------------------------------------------
//			////필수 체크
//			////----------------------------------------------------------------------------------
//			//    If Trim(oDS_PH_PY000A.GetValue("Code", 0)) = "" Then
//			//        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
//			//        oForm.Items("Code").CLICK ct_Regular
//			//        Exit Function
//			//    End If

//			oMat1.FlushToDataSource();
//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY000B.Size > 1)
//				oDS_PH_PY000B.RemoveRecord((oDS_PH_PY000B.Size - 1));
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;

//			//    Set vMatrix = oForm.Items("Mat1").Specific
//			//    vMatrix.Columns("UserCode").TitleObject.Sortable = True
//			//    vMatrix.Columns("UserCode").TitleObject.Sort gst_Ascending
//			//    vMatrix.Columns("UserCode").TitleObject.Sortable = False







//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY000_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY000_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY000A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY000A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY000_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY000_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY000_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}



//		private void PH_PY000_Print_Report01()
//		{

//			string DocNum = null;
//			short ErrNum = 0;
//			string WinTitle = null;
//			string ReportName = null;
//			string sQry = null;

//			string BPLID = null;
//			string ItmBsort = null;
//			string DocDate = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto PH_PY000_Print_Report01_Error;
//			}

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//			WinTitle = "[S142] 발주서";
//			ReportName = "S142_1.rpt";
//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "EXEC PH_PY000_1 '" + oForm.Items.Item("8").Specific.Value + "'";
//			MDC_Globals.gRpt_Formula = new string[2];
//			MDC_Globals.gRpt_Formula_Value = new string[2];
//			MDC_Globals.gRpt_SRptSqry = new string[2];
//			MDC_Globals.gRpt_SRptName = new string[2];
//			MDC_Globals.gRpt_SFormula = new string[2, 2];
//			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//			/// Formula 수식필드

//			/// SubReport


//			/// Procedure 실행"
//			sQry = "EXEC [PS_PP820_01] '" + BPLID + "', '" + ItmBsort + "', '" + DocDate + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				}
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("조회된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			return;
//			PH_PY000_Print_Report01_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY000_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//    }
//}
