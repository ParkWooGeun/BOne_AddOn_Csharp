using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여소급집계처리
    /// </summary>
    internal class PH_PY120 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY120;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY120.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY120_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY120");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY120_CreateItems();
                PH_PY120_FormItemEnabled();
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
        private void PH_PY120_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY120");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY120");
                oDS_PH_PY120 = oForm.DataSources.DataTables.Item("PH_PY120");

                //테이블이 없는경우 데이터셋(Grid)
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("지급구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("총지급액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("총공제액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add("실지급액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //귀속년월
                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
                oForm.DataSources.UserDataSources.Item("YM").ValueEx = DateTime.Now.ToString("yyyyMM");

                //대상기간
                oForm.DataSources.UserDataSources.Add("YMFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YMFrom").Specific.DataBind.SetBound(true, "", "YMFrom");

                oForm.DataSources.UserDataSources.Add("YMTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("YMTo").Specific.DataBind.SetBound(true, "", "YMTo");

                //지급일자
                oForm.DataSources.UserDataSources.Add("JIGBIL", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("JIGBIL").Specific.DataBind.SetBound(true, "", "JIGBIL");

                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY120_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY120_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY120_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                    ////3
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                    ////4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                    ////7
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                    ////8
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:                    ////9
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                    ////12
                //    break;


                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                    ////16
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    ////18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    ////19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                    ////20
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    ////22
                //    /break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                    ////23
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                    ////37
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                    ////38
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag:                    ////39
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
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Btn1")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("월별로 계산된 소급자료를 해당월, 지급일자로 소급급여처리를 진행 하시겠습니까?", 2, "Yes", "No") == 2)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PH_PY120_DataSave();
                        PH_PY120_DataFind();
                    }
                    if (pVal.ItemUID == "Btn_Search")
                    {
                         PH_PY120_DataFind();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY120_DataValidCheck()
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
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY120_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY120_DataFind()
        {
            string sQry;
            int i;

            string[] COLNAM = new string[8];
            string CLTCOD;
            string YM;
            string YMFrom;
            string YMTo;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                YMFrom = oForm.Items.Item("YMFrom").Specific.Value.ToString().Trim();
                YMTo = oForm.Items.Item("YMTo").Specific.Value.ToString().Trim();

                sQry = "Exec PH_PY120_01 '" + CLTCOD + "','" + YMFrom + "','" + YMTo + "'";

                oDS_PH_PY120.ExecuteQuery(sQry);

                //Debug.Print(oDS_PH_PY120.Rows.Count);
                // iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

                //PH_PY120_TitleSetting(iRow);
                COLNAM[0] = "부서";
                COLNAM[1] = "담당";
                COLNAM[2] = "사번";
                COLNAM[3] = "성명";
                COLNAM[4] = "지급구분";
                COLNAM[5] = "총지급액";
                COLNAM[6] = "총공제액";
                COLNAM[7] = "실지급액";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    oGrid1.Columns.Item(i).Editable = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY120_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// DataSave
        /// </summary>
        /// <returns></returns>
        private void PH_PY120_DataSave()
        {
            string sQry;
            string CLTCOD;
            string YM;
            string JIGBIL;
            string YMFrom;
            string YMTo;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                YM = oForm.Items.Item("YM").Specific.Value;
                JIGBIL = oForm.Items.Item("JIGBIL").Specific.Value;
                YMFrom = oForm.Items.Item("YMFrom").Specific.Value;
                YMTo = oForm.Items.Item("YMTo").Specific.Value;
                sQry = "Exec PH_PY111_SOGUBF '" + CLTCOD + "','" + YM + "','" + JIGBIL + "','" + YMFrom + "','" + YMTo + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.MessageBox("소급집계처리가 적용 되었습니다. 급여대장조회 확인 바랍니다.");
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY120_DataSave_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                                        if (Convert.ToDouble(oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0 && Convert.ToDouble(oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0 또는 1만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        oDS_PH_PY120.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                        break;
                                    case "2":
                                        if (Convert.ToDouble(oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0.5 && Convert.ToDouble(oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            oDS_PH_PY120.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0.5 또는 1 만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(oDS_PH_PY120.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) >= 0.5)
                                            {
                                                oDS_PH_PY120.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "56";
                                            }
                                            else
                                            {
                                                oDS_PH_PY120.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY120);
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
