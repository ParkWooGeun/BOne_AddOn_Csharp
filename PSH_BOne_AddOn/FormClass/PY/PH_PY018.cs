using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연봉제휴일교통비체크
    /// </summary>
    internal class PH_PY018 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY018;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY018.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY018_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY018");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY018_CreateItems();
                PH_PY018_FormItemEnabled();
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
        private void PH_PY018_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY018");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY018");
                oDS_PH_PY018 = oForm.DataSources.DataTables.Item("PH_PY018");

                //그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add("일자", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add("요일", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add("휴일근무", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //일자
                oForm.DataSources.UserDataSources.Add("DocDateF", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateF").Specific.DataBind.SetBound(true, "", "DocDateF");
                oForm.DataSources.UserDataSources.Item("DocDateF").Value = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("DocDateT", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateT").Specific.DataBind.SetBound(true, "", "DocDateT");
                oForm.DataSources.UserDataSources.Item("DocDateT").Value = DateTime.Now.ToString("yyyyMMdd");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                //사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY018_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY018_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY018_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY018_DataFind
        /// </summary>
        private void PH_PY018_DataFind()
        {
            short ErrNum = 0;
            int iRow;
            string sQry;
            string CLTCOD;
            string DocDateF;
            string DocDateT;
            string TeamCode;
            string MSTCOD;
            
            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                DocDateF = oForm.Items.Item("DocDateF").Specific.Value.ToString().Trim();
                DocDateT = oForm.Items.Item("DocDateT").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(CLTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "Exec PH_PY018 '" + CLTCOD + "','" + DocDateF + "','" + DocDateT + "','" + TeamCode + "','" + MSTCOD + "'";
                oDS_PH_PY018.ExecuteQuery(sQry);

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY018_TitleSetting(iRow);
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장을 선택하세요, 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY018_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY018_DataSave
        /// </summary>
        private void PH_PY018_DataSave()
        {
            int i;
            string sQry;
            
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                    {
                        sQry = " UPDATE ZPH_PY008 SET Attend = '" + oDS_PH_PY018.Columns.Item("Attend").Cells.Item(i).Value + "'";
                        sQry += " WHERE CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                        sQry += " And PosDate = '" + oDS_PH_PY018.Columns.Item("PosDate").Cells.Item(i).Value.ToString("yyyyMMdd") + "'";
                        sQry += " And MSTCOD = '" + oDS_PH_PY018.Columns.Item("MSTCOD").Cells.Item(i).Value + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                    PSH_Globals.SBO_Application.SetStatusBarMessage("휴일근무가 변경되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY018_DataSave_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// TitleSetting
        /// </summary>
        /// <param name="iRow"></param>
        private void PH_PY018_TitleSetting(int iRow)
        {
            int i;
            string[] COLNAM = new string[5];

            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            
            try
            {
                //그리드 콤보박스
                COLNAM[0] = "일자";
                COLNAM[1] = "사번";
                COLNAM[2] = "성명";
                COLNAM[3] = "요일";
                COLNAM[4] = "휴일근무";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 && i < COLNAM.Length -1)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    else if (i == COLNAM.Length -1)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("Attend");
                        oComboCol.ValidValues.Add("Y", "근무Y");
                        oComboCol.ValidValues.Add("N", "N");
                        oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY018_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboCol);
                oGrid1.AutoResizeColumns();
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        PH_PY018_DataFind();
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        PH_PY018_DataSave();
                    }
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                        {
                            oForm.Freeze(true);
                            for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                            {
                                if (oGrid1.DataTable.GetValue(4, i) == "Y")
                                {
                                    oGrid1.DataTable.Columns.Item(4).Cells.Item(i).Value = "N";
                                }
                                else
                                {
                                    oGrid1.DataTable.Columns.Item(4).Cells.Item(i).Value = "Y";
                                }
                            }
                            oForm.Freeze(false);
                        }
                        else if (pVal.BeforeAction == false)
                        {
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        if (pVal.ItemUID == "CLTCOD")
                        {
                            //기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }
                            sQry = " SELECT '%' AS [Code], '전체' AS [Name], -1 AS [Seq] UNION ALL  ";
                            sQry += "SELECT U_Code AS [Code], U_CodeNm AS [Name],  U_Seq AS [Seq] FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry += " ORDER BY Seq";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");
                            oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("TeamCode").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "Grid01")
                        {
                        }
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("MSTNAM").Specific.Value = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY018);
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
    }
}