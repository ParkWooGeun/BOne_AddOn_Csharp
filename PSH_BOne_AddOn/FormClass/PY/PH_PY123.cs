using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 가압류등록
    /// </summary>
    internal class PH_PY123 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY123;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY123.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY123_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY123");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY123_CreateItems();
                PH_PY123_FormItemEnabled();
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
        private void PH_PY123_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY123");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY123");
                oDS_PH_PY123 = oForm.DataSources.DataTables.Item("PH_PY123");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("일자", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("압류금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("급여구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("변제금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY123").Columns.Add("비고", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 사번(조회용)
                oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

                // 성명(조회용)
                oForm.DataSources.UserDataSources.Add("SFullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SFullName").Specific.DataBind.SetBound(true, "", "SFullName");

                // 일자
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
                oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 구분
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "가압류등록");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("2", "변제등록");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("3", "종결처리");
                oForm.Items.Item("Gubun").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Gubun").DisplayDesc = true;

                // 가압류금액
                oForm.DataSources.UserDataSources.Add("GaapAmt", SAPbouiCOM.BoDataType.dt_SUM, 10);
                oForm.Items.Item("GaapAmt").Specific.DataBind.SetBound(true, "", "GaapAmt");

                // 급여구분
                oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("JOBTYP").Specific.DataBind.SetBound(true, "", "JOBTYP");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("", "");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                // 변제금액
                oForm.DataSources.UserDataSources.Add("BenjAmt", SAPbouiCOM.BoDataType.dt_SUM, 10);
                oForm.Items.Item("BenjAmt").Specific.DataBind.SetBound(true, "", "BenjAmt");

                // 비고
                oForm.DataSources.UserDataSources.Add("Comments", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("Comments").Specific.DataBind.SetBound(true, "", "Comments");
                                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY123_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY123_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY123_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY123);
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
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY123_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        PH_PY123_SAVE();
                        PH_PY123_DataFind();
                    } 
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        PH_PY123_Delete();
                        PH_PY123_DataFind();
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
            string sQry = string.Empty;
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
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                            case "SMSTCOD":
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("SMSTCOD").Specific.VALUE.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("SFullName").Specific.VALUE = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                                break;
                            case "MSTCOD":
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                string Param01 = string.Empty;
                string Param02 = string.Empty;
                string Param03 = string.Empty;
                string Param04 = string.Empty;

                string sQry = string.Empty;
                SAPbobsCOM.Recordset oRecordSet = null;
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);
                            Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            Param02 = oDS_PH_PY123.Columns.Item("DocDate").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY123.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY123.Columns.Item("Gubun").Cells.Item(pVal.Row).Value;

                            if (string.IsNullOrEmpty(Param02))
                            {
                                oForm.Items.Item("Gubun").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("JOBTYP").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("MSTCOD").Specific.VALUE = "";
                                oForm.Items.Item("FullName").Specific.VALUE = "";
                                oForm.DataSources.UserDataSources.Item("DocDate").Value = "";
                                oForm.DataSources.UserDataSources.Item("Comments").Value = "";
                                oForm.DataSources.UserDataSources.Item("GaapAmt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("BenjAmt").Value = "0";
                            }
                            else
                            { 
                            sQry = "EXEC PH_PY123_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
                            oRecordSet.DoQuery(sQry);

                                if ((oRecordSet.RecordCount == 0))
                                {
                                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                                }
                                else
                                {
                                    oForm.DataSources.UserDataSources.Item("DocDate").Value = oRecordSet.Fields.Item("DocDate").Value;
                                    oForm.Items.Item("Gubun").Specific.Select(oRecordSet.Fields.Item("Gubun").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.Items.Item("JOBTYP").Specific.Select(oRecordSet.Fields.Item("JOBTYP").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("MSTCOD").Value;
                                    oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                    oForm.DataSources.UserDataSources.Item("GaapAmt").Value = oRecordSet.Fields.Item("GaapAmt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("BenjAmt").Value = oRecordSet.Fields.Item("BenjAmt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("Comments").Value = oRecordSet.Fields.Item("Comments").Value;
                                    oForm.Update();
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY123_DataFind
        /// </summary>
        private void PH_PY123_DataFind()
        {
            short ErrNum = 0;
            int iRow = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string SMSTCOD = string.Empty;
            
            string[] COLNAM = new string[5];

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                SMSTCOD = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim();
                
                if (string.IsNullOrEmpty(SMSTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                sQry = "Exec PH_PY123_01 '" + CLTCOD + "','" + SMSTCOD + "'";
                oDS_PH_PY123.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY123_TitleSetting(ref iRow);
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사원코드를 입력 하세요, 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY123_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY123_SAVE
        /// </summary>
        private void PH_PY123_SAVE()
        {
            // 데이타 저장
            short ErrNum = 0;
            string sQry = string.Empty;
            string JOBTYP = string.Empty;
            string MSTCOD = string.Empty;
            string CLTCOD = string.Empty;
            string DocDate = string.Empty;
            string Gubun = string.Empty;
            string Comments = string.Empty; ;
            double GaapAmt = 0;
            double BenjAmt = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
            
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                Gubun = oForm.Items.Item("Gubun").Specific.VALUE.ToString().Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.VALUE.ToString().Trim();
                GaapAmt = Convert.ToDouble(oForm.Items.Item("GaapAmt").Specific.VALUE);
                BenjAmt = Convert.ToDouble(oForm.Items.Item("BenjAmt").Specific.VALUE);
                Comments = oForm.Items.Item("Comments").Specific.VALUE.ToString().Trim();
                
                if (JOBTYP.Trim().Length != 0)
                {                
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(Gubun))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                sQry = "Select Cnt = Count(*) From ZPH_PY123 Where CLTCOD = '" + CLTCOD + "' AND MSTCOD = '" + MSTCOD + "'";
                sQry = sQry + " And DocDate = '" + DocDate + "' and Gubun = '" + Gubun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    //신규
                    sQry = "INSERT INTO ZPH_PY123";
                    sQry = sQry + " (";
                    sQry = sQry + " CLTCOD,";
                    sQry = sQry + " DocDate,";
                    sQry = sQry + " MSTCOD,";
                    sQry = sQry + " Gubun,";
                    sQry = sQry + " JOBTYP,";
                    sQry = sQry + " GaapAmt,";
                    sQry = sQry + " BenjAmt,";
                    sQry = sQry + " Comments";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";
                    sQry = sQry + "'" + CLTCOD + "',";
                    sQry = sQry + "'" + DocDate + "',";
                    sQry = sQry + "'" + MSTCOD + "',";
                    sQry = sQry + "'" + Gubun + "',";
                    sQry = sQry + "'" + JOBTYP + "',";
                    sQry = sQry + GaapAmt + ",";
                    sQry = sQry + BenjAmt + ",";
                    sQry = sQry + "'" + Comments + "'";
                    sQry = sQry + ")";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                else
                {
                    //수정
                    sQry = "Update ZPH_PY123";
                    sQry = sQry + " Set DocDate = '" + DocDate + "',";
                    sQry = sQry + " Gubun = '" + Gubun + "',";
                    sQry = sQry + " MSTCOD = '" + MSTCOD + "',";
                    sQry = sQry + " GaapAmt = " + GaapAmt + ",";
                    sQry = sQry + " BenjAmt = " + BenjAmt + ",";
                    sQry = sQry + " Comments = '" + Comments + "'";
                    sQry = sQry + " Where CLTCOD = '" + CLTCOD + "'";
                    sQry = sQry + " And DocDate = '" + DocDate + "'";
                    sQry = sQry + " And MSTCOD = '" + MSTCOD + "'";
                    sQry = sQry + " And Gubun = '" + Gubun + "'";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("급여계산 된 자료는 수정할 수 없습니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("구분은 필수 입력사항입니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY123_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY123_Delete
        /// </summary>
        private void PH_PY123_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string JOBTYP = string.Empty;
            string MSTCOD = string.Empty;
            string CLTCOD = string.Empty;
            string DocDate = string.Empty;
            string Gubun = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                Gubun = oForm.Items.Item("Gubun").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.VALUE.ToString().Trim();

                if (JOBTYP.Trim().Length != 0)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY123.Rows.Count > 0)
                    {
                        sQry = "Delete From ZPH_PY123 Where CLTCOD = '" + CLTCOD + "' AND  DocDate = '" + DocDate + "' And MSTCOD = '" + MSTCOD + "' And Gubun = '" + Gubun + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("급여계산 된 자료는 삭제할 수 없습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY123_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY123_TitleSetting
        /// </summary>
        private void PH_PY123_TitleSetting(ref int iRow)
        {
            int i = 0;
            string[] COLNAM = new string[9];

            try
            {
                //그리드 콤보박스
                COLNAM[0] = "사업장";
                COLNAM[1] = "사번";
                COLNAM[2] = "성명";
                COLNAM[3] = "일자";
                COLNAM[4] = "구분";
                COLNAM[5] = "압류금액";
                COLNAM[6] = "급여구분";
                COLNAM[7] = "변제금액";
                COLNAM[8] = "비고";

                for (i = 0; i <= Information.UBound(COLNAM) ; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i <= Information.UBound(COLNAM))
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY123_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

    }
}

