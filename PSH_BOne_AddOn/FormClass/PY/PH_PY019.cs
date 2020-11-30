using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 반변경등록
    /// </summary>
    internal class PH_PY019 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY019;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY019.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY019_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY019");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY019_CreateItems();
                PH_PY019_FormItemEnabled();
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
        private void PH_PY019_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY019");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY019");
                oDS_PH_PY019 = oForm.DataSources.DataTables.Item("PH_PY019");

                //그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY019").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY019").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY019").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY019").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY019").Columns.Add("반", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                oForm.Items.Item("TeamCode").DisplayDesc = true;
                
                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                //반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");
                oForm.Items.Item("ClsCode").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY019_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY019_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY019_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY019_DataFind
        /// </summary>
        private void PH_PY019_DataFind()
        {
            short ErrNum = 0;
            int iRow;
            string sQry;
            string CLTCOD;
            string TeamCode;
            string RspCode;
            string ClsCode;

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(TeamCode))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(RspCode))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                sQry = "Exec PH_PY019 '" + CLTCOD + "','" + TeamCode + "','" + RspCode + "','" + ClsCode + "'";
                oDS_PH_PY019.ExecuteQuery(sQry);

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY019_TitleSetting(iRow);
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("부서를 선택하세요, 확인바랍니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("담당을 선택하세요, 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY019_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY019_DataSave
        /// </summary>
        private void PH_PY019_DataSave()
        {
            string sQry;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                    {
                        sQry = " UPDATE [@PH_PY001A] SET U_ClsCode = '" + oDS_PH_PY019.Columns.Item("ClsCode").Cells.Item(i).Value + "'";
                        sQry += ",                 U_UpdtProg = 'PH_PY019_Y'";
                        sQry += ",                 U_UserSign2 = '" + PSH_Globals.oCompany.UserSignature.ToString() + "'";
                        sQry += " WHERE Code = '" + oDS_PH_PY019.Columns.Item("MSTCOD").Cells.Item(i).Value + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                    PSH_Globals.SBO_Application.SetStatusBarMessage("반이 변경되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {

                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY019_DataSave_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

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
        private void PH_PY019_TitleSetting(int iRow)
        {
            int i;
            int j;
            string sQry;
            string[] COLNAM = new string[5];

            SAPbouiCOM.ComboBoxColumn oComboCol;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //그리드 콤보박스
                COLNAM[0] = "사번";
                COLNAM[1] = "성명";
                COLNAM[2] = "부서";
                COLNAM[3] = "담당";
                COLNAM[4] = "반";

                for (i = 0; i <= COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 && i < COLNAM.Length)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    else if (i == COLNAM.Length)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("ClsCode");

                        sQry = " SELECT  U_Code AS [Code], U_CodeNm AS [Name],  U_Seq AS [Seq] FROM [@PS_HR200L] ";
                        sQry += " WHERE       Code = '9' ";
                        sQry += "             AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                        sQry += " ORDER BY    Seq";

                        oRecordSet.DoQuery(sQry);
                        if (oRecordSet.RecordCount > 0)
                        {
                            oComboCol.ValidValues.Add("", "");
                            for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                            {
                                oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                oRecordSet.MoveNext();
                            }
                        }
                        oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY019_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        PH_PY019_DataFind();
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        PH_PY019_DataSave();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
                            // 기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry += " ORDER BY U_Seq";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");

                            oForm.Items.Item("TeamCode").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "TeamCode")
                        {
                            // 담당 (사업장에 따른 담당변경)
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '2' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "' And U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry += " Order By U_Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");
                            oForm.Items.Item("RspCode").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "RspCode")
                        {
                            // 반(사업장에 따른 반변경)
                            if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT  U_Code AS [Code], U_CodeNm AS [Name],  U_Seq AS [Seq] FROM [@PS_HR200L] ";
                            sQry += " WHERE       Code = '9' ";
                            sQry += "             AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_Char2 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "N");
                            oForm.Items.Item("ClsCode").DisplayDesc = true;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY019);
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
    }
}