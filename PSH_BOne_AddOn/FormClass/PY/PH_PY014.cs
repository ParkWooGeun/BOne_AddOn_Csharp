using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 위해일수수정
    /// </summary>
    internal class PH_PY014 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY014;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY014.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY014_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY014");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY014_CreateItems();
                PH_PY014_FormItemEnabled();
                PSH_Globals.ExecuteEventFilter(typeof(PH_PY014));
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
        /// <returns></returns>
        private void PH_PY014_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY014");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY014");
                oDS_PH_PY014 = oForm.DataSources.DataTables.Item("PH_PY014");

                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("일자", SAPbouiCOM.BoFieldsType.ft_Date);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("요일", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("근태구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("위해일수", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY014").Columns.Add("위해코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                
                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY014_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY014_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY014_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private bool PH_PY014_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY014_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PH_PY014_DataFind()
        {
            int iRow;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sQry = "Exec PH_PY014 '";
                sQry += oForm.Items.Item("CLTCOD").Specific.Value + "','";
                sQry += oForm.Items.Item("YM").Specific.Value + "','";
                sQry += oForm.Items.Item("MSTCOD").Specific.Value + "'";
                oDS_PH_PY014.ExecuteQuery(sQry);

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

                PH_PY014_TitleSetting(iRow);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY014_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// DataSave
        /// </summary>
        /// <returns></returns>
        private bool PH_PY014_DataSave()
        {
            bool functionReturnValue = false;
            int i;
            string ShiftDat;
            string sQry;
            string CLTCOD;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();

                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                    {
                        ShiftDat = oDS_PH_PY014.Columns.Item("PosDate").Cells.Item(i).Value.ToString("yyyyMMdd");

                        sQry = "  UPDATE ZPH_PY008 SET DangerNu = '" + oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(i).Value + "',";
                        sQry += " DangerCD = '" + oDS_PH_PY014.Columns.Item("DangerCD").Cells.Item(i).Value + "'";
                        sQry += " WHERE CLTCOD = '" + CLTCOD + "'";
                        sQry += " And PosDate = '" + ShiftDat + "'";
                        sQry += " And MSTCOD = '" + oDS_PH_PY014.Columns.Item("MSTCOD").Cells.Item(i).Value + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                    PSH_Globals.SBO_Application.MessageBox("위해일수(등급)가 변경되었습니다.");

                    functionReturnValue = true;
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox("데이터가 존재하지 않습니다.");
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY014_DataSave_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 그리드 타이블 변경
        /// </summary>
        /// <param name="iRow"></param>
        private void PH_PY014_TitleSetting(int iRow)
        {
            int i;
            int j;
            string sQry;
            string[] COLNAM = new string[9];
            string CLTCOD;

            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "일자";
                COLNAM[1] = "요일";
                COLNAM[2] = "근태구분";
                COLNAM[3] = "부서";
                COLNAM[4] = "담당";
                COLNAM[5] = "사번";
                COLNAM[6] = "성명";
                COLNAM[7] = "위해일수";
                COLNAM[8] = "위해코드";

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    switch (COLNAM[i])
                    {
                        case "근태구분":
                            oGrid1.Columns.Item(i).Editable = false;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("WorkType");

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P221' AND U_UseYN= 'Y' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                    oRecordSet.MoveNext();
                                }
                            }
                            oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;

                        case "위해일수":
                            oGrid1.Columns.Item(i).Editable = true;
                            oGrid1.Columns.Item(i).RightJustified = true;
                            break;

                        case "위해코드":
                            oGrid1.Columns.Item(i).Editable = true;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("DangerCD");

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P220' AND U_UseYN= 'Y' And U_Char2 = '" + CLTCOD + "' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            oComboCol.ValidValues.Add("", "");
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                    oRecordSet.MoveNext();
                                }
                            }
                            oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                        default:
                            oGrid1.Columns.Item(i).Editable = false;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY014_TitleSetting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboCol);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    break;
                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        if (PH_PY014_DataValidCheck() == true)
                        {
                            PH_PY014_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        if (PH_PY014_DataSave() == false)
                        {
                            BubbleEvent = false;
                        }
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
        /// ITEM_PRESSED 이벤트
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                sQry = "SELECT U_FullName from [@PH_PY001A] Where Code = '" + oForm.Items.Item("MSTCOD").Specific.Value + "'";
                                oRecordSet.DoQuery(sQry);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                }
                                break;

                            case "Grid01":
                                switch (oForm.Items.Item("CLTCOD").Specific.Value.Trim())
                                {
                                    case "1":
                                        if (Convert.ToDouble(oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0 && Convert.ToDouble(oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0 또는 1만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        oDS_PH_PY014.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                        break;
                                    case "2":
                                        if (Convert.ToDouble(oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0.5 && Convert.ToDouble(oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            oDS_PH_PY014.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0.5 또는 1 만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(oDS_PH_PY014.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) >= 0.5)
                                            {
                                                oDS_PH_PY014.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "56";
                                            }
                                            else
                                            {
                                                oDS_PH_PY014.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                            }
                                        }
                                        break;
                                }
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