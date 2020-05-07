using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.VisualBasic;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여변동자료등록
    /// </summary>
    internal class PH_PY109 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        public SAPbouiCOM.Matrix oMat2;
        private SAPbouiCOM.DBDataSource oDS_PH_PY109A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY109B;
        private SAPbouiCOM.DBDataSource oDS_PH_PY109Z;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY109.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PH_PY109_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY109");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY109_CreateItems();
                PH_PY109_EnableMenus();
                PH_PY109_SetDocument(oFromDocEntry01);
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
        private void PH_PY109_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY109A = oForm.DataSources.DBDataSources.Item("@PH_PY109A"); // 헤더
                oDS_PH_PY109B = oForm.DataSources.DBDataSources.Item("@PH_PY109B"); // 라인
                oDS_PH_PY109Z = oForm.DataSources.DBDataSources.Item("@PH_PY109Z"); // 라인

                oMat1 = oForm.Items.Item("Mat1").Specific;
                oMat2 = oForm.Items.Item("Mat2").Specific;

                // 사업장
             //   oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
             //   oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                // 접속자에 따른 권한별 사업장 콤보박스세팅
                //dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                // 귀속년월
             //   oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
             //   oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
                oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

                // 지급종류
                oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                // 지급구분
                oForm.DataSources.UserDataSources.Add("JOBGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                // 지급대상
                oForm.DataSources.UserDataSources.Add("JOBTRG", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBTRG").Specific, "");
                oForm.Items.Item("JOBTRG").DisplayDesc = true;

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                oForm.DataSources.UserDataSources.Add("S_Amt01", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt01").Specific.DataBind.SetBound(true, "", "S_Amt01");

                oForm.DataSources.UserDataSources.Add("S_Amt02", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt02").Specific.DataBind.SetBound(true, "", "S_Amt02");

                oForm.DataSources.UserDataSources.Add("S_Amt03", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt03").Specific.DataBind.SetBound(true, "", "S_Amt03");

                oForm.DataSources.UserDataSources.Add("S_Amt04", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt04").Specific.DataBind.SetBound(true, "", "S_Amt04");

                oForm.DataSources.UserDataSources.Add("S_Amt05", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt05").Specific.DataBind.SetBound(true, "", "S_Amt05");

                oForm.DataSources.UserDataSources.Add("S_Amt06", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt06").Specific.DataBind.SetBound(true, "", "S_Amt06");

                oForm.DataSources.UserDataSources.Add("S_Amt07", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt07").Specific.DataBind.SetBound(true, "", "S_Amt07");

                oForm.DataSources.UserDataSources.Add("S_Amt08", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt08").Specific.DataBind.SetBound(true, "", "S_Amt08");

                oForm.DataSources.UserDataSources.Add("S_Amt09", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt09").Specific.DataBind.SetBound(true, "", "S_Amt09");

                oForm.DataSources.UserDataSources.Add("S_Amt10", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt10").Specific.DataBind.SetBound(true, "", "S_Amt10");

                oForm.DataSources.UserDataSources.Add("S_Amt11", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt11").Specific.DataBind.SetBound(true, "", "S_Amt11");

                oForm.DataSources.UserDataSources.Add("S_Amt12", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Amt12").Specific.DataBind.SetBound(true, "", "S_Amt12");

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY109_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);  // 제거
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", true);  // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY109_SetDocument
        /// </summary>
        private void PH_PY109_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY109_FormItemEnabled();
                    PH_PY109_AddMatrixRow();
                    PH_PY109_TitleSetting_Matrix();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY109_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY109_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("JOBTYP").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("JOBTRG").Enabled = true;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    // 귀속년월
                    oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

                    oForm.EnableMenu("1281", true);  // 문서찾기
                    oForm.EnableMenu("1282", false); // 문서추가

                    PH_PY109_AddMatrixRow();
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("JOBTYP").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("JOBTRG").Enabled = true;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false);  // 문서찾기
                    oForm.EnableMenu("1282", true);   // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("JOBTYP").Enabled = false;
                    oForm.Items.Item("JOBGBN").Enabled = false;
                    oForm.Items.Item("JOBTRG").Enabled = false;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true);   // 문서찾기
                    oForm.EnableMenu("1282", true);   // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY109_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                //    '//[Mat1 용]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY109B.GetValue("U_MSTCOD", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY109B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY109B.InsertRecord((oRow));
                        }
                        oDS_PH_PY109B.Offset = oRow;
                        oDS_PH_PY109B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY109B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY109B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY109B.SetValue("U_DPTCOD", oRow, "");
                        oDS_PH_PY109B.SetValue("U_DPTNAM", oRow, "");
                        oDS_PH_PY109B.SetValue("U_AMT01", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT02", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT03", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT04", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT05", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT06", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT07", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT08", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT09", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT10", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT11", oRow, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT12", oRow, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY109B.Offset = oRow - 1;
                        oDS_PH_PY109B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY109B.SetValue("U_MSTCOD", oRow - 1, "");
                        oDS_PH_PY109B.SetValue("U_MSTNAM", oRow - 1, "");
                        oDS_PH_PY109B.SetValue("U_DPTCOD", oRow - 1, "");
                        oDS_PH_PY109B.SetValue("U_DPTNAM", oRow - 1, "");
                        oDS_PH_PY109B.SetValue("U_AMT01", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT02", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT03", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT04", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT05", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT06", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT07", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT08", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT09", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT10", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT11", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY109B.SetValue("U_AMT12", oRow - 1, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY109B.Offset = oRow;
                    oDS_PH_PY109B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY109B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY109B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY109B.SetValue("U_DPTCOD", oRow, "");
                    oDS_PH_PY109B.SetValue("U_DPTNAM", oRow, "");
                    oDS_PH_PY109B.SetValue("U_AMT01", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT02", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT03", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT04", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT05", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT06", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT07", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT08", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT09", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT10", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT11", oRow, Convert.ToString(0));
                    oDS_PH_PY109B.SetValue("U_AMT12", oRow, Convert.ToString(0));
                    oMat1.LoadFromDataSource();
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            int i = 0;
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY109A", "Code");
                           	// 접속자 권한에 따른 사업장 보기
                            break;

                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY109_FormItemEnabled();
                            break;


                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": // 문서찾기
                            PH_PY109_FormItemEnabled();
                            Sum_Display();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": // 문서추가
                            PH_PY109_FormItemEnabled();
                            Sum_Display();
                            break;

                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY109_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제
                            // [MAT1 용]
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY109B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY109B.GetValue("U_MSTCOD", i).Trim()))
                                    {
                                        oDS_PH_PY109B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY109B.Size; i++)
                                {
                                    oDS_PH_PY109B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
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
            int i = 0;
            string FullName = string.Empty;
            string FindYN = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY109_DataValidCheck(pVal.ItemUID) == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Btn_Set")
                    {
                        if (PH_PY109_DataValidCheck(pVal.ItemUID) == false)
                        {
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY109_FormItemEnabled();
                            PH_PY109_TitleSetting_Matrix();
                        }
                    }
                    if (pVal.ItemUID == "Btn_Upload")
                    {
                        oMat1.FlushToDataSource();

                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY109_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();

                        PH_PY109_AddMatrixRow();

                        Sum_Display();
                    }
                    if (pVal.ItemUID == "Btn_Set")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY109_LoadData_SudangGongje();
                        }
                    }

                    if (pVal.ItemUID == "Search")
                    {
                        FindYN = "N";
                        FullName = oForm.Items.Item("FullName").Specific.VALUE;
                        for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                        {
                            if (oMat1.Columns.Item("MSTNAM").Cells.Item(i).Specific.VALUE.Trim() == FullName)
                            {
                                FindYN = "Y";
                                oMat1.SelectRow(i, true, false);
                            }
                        }
                        if (FindYN == "N")
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("찾는 사원이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                    case "Mat2":
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
                        case "Mat1":
                        case "Mat2":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);

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

                    PH_PY109_FormItemEnabled();
                    PH_PY109_AddMatrixRow();
                    PH_PY109_TitleSetting_Matrix();

                    Sum_Display();

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
        /// Raise_EVENT_KEY_DOWN
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true & pVal.ItemUID == "Mat1" & pVal.ColUID == "MSTCOD" & pVal.CharPressed == 9)
                {
                    if (string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE.Trim()))
                    {
                        oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                        BubbleEvent = false;
                    }
                    else
                    {
                        if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", Convert.ToString(Convert.ToDouble("'") + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE + Convert.ToDouble("'")),"") == true)
                        {
                            oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "MSTCOD")
                        {
                            if (!string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE.Trim()))
                            {
                                oMat1.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.VALUE = dataHelpClass.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE.Trim() + "'","");
                                oMat1.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.VALUE = dataHelpClass.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.VALUE.Trim() + "'","");
                                oMat1.Columns.Item("DPTNAM").Cells.Item(pVal.Row).Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat1.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.VALUE + "'", " AND Code = '1'");
                            }
                            PH_PY109_AddMatrixRow();
                            oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "MSTCOD")
                        {
                            PH_PY109_AddMatrixRow();
                            oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY109A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY109B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY109Z);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat2);
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
        /// Raise_EVENT_CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat1")
                    {
                        // 프로젝트 코드
                        if (pVal.ColUID == "MSTCOD")
                        {
                            dataHelpClass.PSH_CF_DBDatasourceReturn( pVal, pVal.FormUID, "@PH_PY109B", "U_MSTCOD,U_MSTNAM,U_DPTCOD,U_DPTNAM", "Mat1", (short) pVal.Row, "", "", "");
                            PH_PY109_AddMatrixRow();
                            oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oMat1.AutoResizeColumns();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY109_DataValidCheck(string ItemUID)
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY109A.GetValue("U_CLTCOD", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY109A.GetValue("U_YM", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("귀속년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY109A.GetValue("U_JOBTYP", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급종류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JOBTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY109A.GetValue("U_JOBGBN", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY109A.GetValue("U_JOBTRG", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급대상은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("JOBTRG").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }


                oDS_PH_PY109A.SetValue("Code", 0, oDS_PH_PY109A.GetValue("U_CLTCOD", 0).Trim() + codeHelpClass.Right(oDS_PH_PY109A.GetValue("U_YM", 0).Trim(), 4) + oDS_PH_PY109A.GetValue("U_JOBTYP", 0).Trim() + oDS_PH_PY109A.GetValue("U_JOBGBN", 0).Trim() + oDS_PH_PY109A.GetValue("U_JOBTRG", 0).Trim());
                oDS_PH_PY109A.SetValue("Name", 0, oDS_PH_PY109A.GetValue("COde", 0).Trim());

                if (ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" + oDS_PH_PY109A.GetValue("COde", 0).Trim() + "'","")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }
                else if (ItemUID == "Btn_Set")
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" + oDS_PH_PY109A.GetValue("COde", 0).Trim() + "'", "")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }

                oMat1.FlushToDataSource();

                if (ItemUID == "1")
                {
                    // Matrix 마지막 행 삭제(DB 저장시)
                    if (oDS_PH_PY109B.Size > 1)
                    {
                        oDS_PH_PY109B.RemoveRecord((oDS_PH_PY109B.Size - 1));
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        functionReturnValue = false;
                        return functionReturnValue;
                    }
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY109_DataValidCheckError:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// DataFind : 자료 조회
        /// </summary>
        private void PH_PY109_SaveData()
        {
            int i = 0;
            string sQry = string.Empty;

            string Code = string.Empty;
            short Sequence = 0;
            string PDCode = string.Empty;
            string PDName = string.Empty;
            string bPDCode = string.Empty;
            string bPDName = string.Empty;
            double Amt = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();

                for (i = 0; i <= oDS_PH_PY109Z.Size - 1; i++)
                {
                    Code = oDS_PH_PY109Z.GetValue("U_ColReg01", i).ToString().Trim();
                    Sequence = Convert.ToInt16(oDS_PH_PY109Z.GetValue("U_ColNum01", i).ToString().Trim());
                    PDCode = oDS_PH_PY109Z.GetValue("U_ColReg03", i).ToString().Trim();
                    PDName = oDS_PH_PY109Z.GetValue("U_ColReg04", i).ToString().Trim();
                    bPDCode = oDS_PH_PY109Z.GetValue("U_ColReg05", i).ToString().Trim();
                    bPDName = oDS_PH_PY109Z.GetValue("U_ColReg06", i).ToString().Trim();
                    Amt = Convert.ToDouble(oDS_PH_PY109Z.GetValue("U_ColSum01", i).ToString().Trim());

                    ////문서번호가 있고
                    if (!string.IsNullOrEmpty(Code))
                    {

                        ////수당코드가 수정이 되었으면 Update대상
                        if (PDCode != bPDCode)
                        {
                            if (Amt == 0)
                            {
                                sQry = "Update [@PH_PY109Z] Set U_PDCode = '" + PDCode + "' , U_PDName = '" + PDName + "'";
                                sQry = sQry + " Where Code = '" + Code + "' And U_Sequence = " + Sequence + "";
                                sQry = sQry + " And U_PDCode = '" + bPDCode + "' And U_PDName = '" + bPDName + "'";

                                oRecordSet.DoQuery(sQry);
                                PSH_Globals.SBO_Application.MessageBox("저장되었습니다. 급여변동자료 등록에서 확인바랍니다.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY109_SaveData_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// DataFind : 자료 조회
        /// </summary>
        private void PH_PY109_LoadData()
        {
            int i = 0;
            string sQry = string.Empty;

            string CLTCOD = string.Empty;
            string YM = string.Empty;
            string JOBTYP = string.Empty;
            string JOBGBN = string.Empty;
            string JOBTRG = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.VALUE.ToString().Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.VALUE.ToString().Trim();
                JOBGBN = oForm.Items.Item("JOBGBN").Specific.VALUE.ToString().Trim();
                JOBTRG = oForm.Items.Item("JOBTRG").Specific.VALUE.ToString().Trim();

                //// 수당, 공제 테이블 고정:V, 상여:Y 인 값을 임시테이블에 넣는다
                sQry = "EXEC PH_PY109_01 '" + CLTCOD + "' , '" + YM + "' , '" + JOBTYP + "' , '" + JOBGBN + "' , '" + JOBTRG + "'";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if ((oRecordSet.RecordCount == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                }
                else
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        if (i != 0)
                        {
                            oDS_PH_PY109Z.InsertRecord((i));
                        }
                        oDS_PH_PY109Z.Offset = i;
                        oDS_PH_PY109Z.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Code").Value);
                        oDS_PH_PY109Z.SetValue("U_ColNum01", i, oRecordSet.Fields.Item("Sequence").Value);
                        oDS_PH_PY109Z.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("PayDud").Value);
                        oDS_PH_PY109Z.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("PDCode").Value);
                        oDS_PH_PY109Z.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("PDName").Value);
                        oDS_PH_PY109Z.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value);
                        oDS_PH_PY109Z.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("PDCode").Value);
                        oDS_PH_PY109Z.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("PDName").Value);
                        oRecordSet.MoveNext();
                    }
                    oMat1.LoadFromDataSource();
                    oMat1.AutoResizeColumns();
                    oForm.Update();

                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY109_LoadData_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY109_TitleSetting_Matrix
        /// </summary>
        private void PH_PY109_TitleSetting_Matrix()
        {
            int i = 0;
            int iCount = 0;

            try
            {
                oForm.Freeze(true);
                iCount = 0;
                for (i = 5; i <= oMat1.Columns.Count - 1; i++)
                {
                    if (oDS_PH_PY109Z.Size >= 1)
                    {
                        if (string.IsNullOrEmpty(oDS_PH_PY109Z.GetValue("U_PDName", 0).Trim()))
                        {
                            oMat1.Columns.Item(1).Editable = false;
                            oMat1.Columns.Item(i).TitleObject.Caption = "";
                            oMat1.Columns.Item(i).Editable = false;
                        }
                        else
                        {
                            oMat1.Columns.Item(1).Editable = true;
                            if (i > 4 & i <= oDS_PH_PY109Z.Size + 4)
                            {
                                oMat1.Columns.Item(i).TitleObject.Caption = oDS_PH_PY109Z.GetValue("U_PDName", iCount).Trim();
                                oMat1.Columns.Item(i).Editable = true;
                                iCount = iCount + 1;
                            }
                            else
                            {
                                oMat1.Columns.Item(i).TitleObject.Caption = "";
                                oMat1.Columns.Item(i).Editable = false;
                            }
                        }
                    }
                    else
                    {
                        oMat1.Columns.Item(1).Editable = false;
                        oMat1.Columns.Item(i).TitleObject.Caption = "";
                        oMat1.Columns.Item(i).Editable = false;
                    }
                }

                oMat1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY109_TitleSetting_Matrix_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY109_LoadData_SudangGongje
        /// </summary>
        private object PH_PY109_LoadData_SudangGongje()
        {
            object functionReturnValue = null;
            string sQry = string.Empty;
            string sCode = string.Empty;
            string sCLTCOD = string.Empty;
            string sYM = string.Empty;
            string sJOBTYP = string.Empty;
            string sJOBGBN = string.Empty;
            string sJOBTRG = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (oDS_PH_PY109Z.Size > 0)
                {
                    oDS_PH_PY109Z.Clear();
                }

                sCode =   oDS_PH_PY109A.GetValue("Code", 0).Trim();
                sCLTCOD = oDS_PH_PY109A.GetValue("U_CLTCOD", 0).Trim();
                sYM =     oDS_PH_PY109A.GetValue("U_YM", 0).Trim();
                sJOBTYP = oDS_PH_PY109A.GetValue("U_JOBTYP", 0).Trim();
                sJOBGBN = oDS_PH_PY109A.GetValue("U_JOBGBN", 0).Trim();
                sJOBTRG = oDS_PH_PY109A.GetValue("U_JOBTRG", 0).Trim();

                // 수당, 공제 테이블 고정:V, 상여:Y 인 값을 임시테이블에 넣는다
                sQry = "EXEC PH_PY109 '" + oDS_PH_PY109A.GetValue("U_CLTCOD", 0).Trim() + "' , '" + oDS_PH_PY109A.GetValue("U_YM", 0).Trim() + "' , '";
                sQry = sQry + oDS_PH_PY109A.GetValue("U_JOBTYP", 0).Trim() + "' , '" + oDS_PH_PY109A.GetValue("U_JOBGBN", 0).Trim() + "' , '";
                sQry = sQry + oDS_PH_PY109A.GetValue("U_JOBTRG", 0).Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                oForm.Items.Item("CLTCOD").Specific.Select("" + sCLTCOD + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("YM").Specific.VALUE = sYM;
                oForm.Items.Item("JOBTYP").Specific.Select("" + sJOBTYP + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JOBGBN").Specific.Select("" + sJOBGBN + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JOBTRG").Specific.Select("" + sJOBTRG + "", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                return functionReturnValue;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY109_LoadData_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return functionReturnValue;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY109_Excel_Upload 
        /// </summary>
        [STAThread]
        private void PH_PY109_Excel_Upload()
        {
            int i = 0;
            int j = 0;
            int CheckLine = 0;
            int ErrNum = 0;
            int TOTCNT = 0;
            int V_StatusCnt = 0;
            int oProValue = 0;
            int tRow = 0;
            bool CheckYN = false;
            string sPrice = string.Empty;
            string sFile = string.Empty;
            string OneRec = string.Empty;
            string sQry = string.Empty;
            string MSTCOD = string.Empty;

            short columnCount = 3;  // 엑셀 컬럼수
            short columnCount2 = 3; // 엑셀 컬럼수
            int loopCount = 0;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            //원본
            //Microsoft.Office.Interop.Excel.Application xl = default(Microsoft.Office.Interop.Excel.Application);
            //Microsoft.Office.Interop.Excel.Workbook xlwb = default(Microsoft.Office.Interop.Excel.Workbook);
            //Microsoft.Office.Interop.Excel.Worksheet xlsh = default(Microsoft.Office.Interop.Excel.Worksheet);

            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));

            if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                sFile = commonOpenFileDialog.FileName;
            }
            else //Cancel 버튼 클릭
            {
                return;
            }

            //엑셀 Object 연결
            //암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
            Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);  // sFile
            Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
            Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
            Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
            Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.ActiveForm;

                if (oMat1.Columns.Item("MSTCOD").Editable == false)
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(sFile))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                else
                {
                    if (codeHelpClass.Right(sFile,  3) != "xls" & codeHelpClass.Right(sFile,  4) != "xlsx")
                    {
                        ErrNum = 3;
                        throw new Exception();
                    }
                }

                if (xlsh.UsedRange.Columns.Count <= 2)
                {
                    ErrNum = 4;
                    throw new Exception();
                }

                Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
                for (loopCount = 1; loopCount <= columnCount2; loopCount++)
                {
                    t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
                }

                if (t[1].Value.ToString().Trim() != "사번")
                {
                    ErrNum = 5;
                    throw new Exception();
                }

                if (t[2].Value.ToString().Trim() != "성명")
                {
                    ErrNum = 6;
                    throw new Exception();
                }

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);
                // 최대값 구하기
                TOTCNT = xlsh.UsedRange.Rows.Count - 1;
                V_StatusCnt = TOTCNT / 50;
                oProValue = 1;
                tRow = 1;

                for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
                    }

                    // 사번 존재 여부 체크
                    if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", "'" + r[1].Value.ToString().Trim() + "'","") == true)
                    {
                        ErrNum = 7;
                        throw new Exception();
                    }
                    else
                    {
                        CheckYN = false;

                        for (j = 0; j <= oDS_PH_PY109B.Size - 1; j++)
                        {
                            if (r[1].Value.ToString().Trim() == oDS_PH_PY109B.GetValue("U_MSTCOD", j).ToString().Trim())
                            {
                                CheckYN = true;
                                CheckLine = j;
                            }
                        }

                        // 마지막행 제거
                        if (string.IsNullOrEmpty(oDS_PH_PY109B.GetValue("U_MSTCOD", oDS_PH_PY109B.Size - 1).ToString().Trim()))
                        {
                            oDS_PH_PY109B.RemoveRecord((oDS_PH_PY109B.Size - 1));
                        }

                        // 사원마스터에서 사번에 대한 정보 가져오기
                        MSTCOD = r[1].Value.ToString().Trim();
                        sQry = "select U_FullName, U_TeamCode, U_CodeNm ";
                        sQry = sQry + " FROM [@PH_PY001A] T0 INNER JOIN [@Ps_HR200L] T1 ON T0.U_Teamcode = T1.U_Code";
                        sQry = sQry + " WHERE T0.Code = '" + MSTCOD + "'";
                        oRecordSet.DoQuery(sQry);

                        // Mat1에 업로드엑셀 사번이 존재 유무 CheckYN
                        if (CheckYN == true)
                        {
                            oDS_PH_PY109B.SetValue("U_MSTNAM", CheckLine, oRecordSet.Fields.Item(0).Value);
                            oDS_PH_PY109B.SetValue("U_DPTCOD", CheckLine, oRecordSet.Fields.Item(1).Value);
                            oDS_PH_PY109B.SetValue("U_DPTNAM", CheckLine, oRecordSet.Fields.Item(2).Value);

                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT01").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT01", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT02").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT02", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT03").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT03", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT04").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT04", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT05").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT05", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT06").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT06", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT07").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT07", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT08").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT08", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT09").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT09", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT10").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT10", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT11").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT11", CheckLine, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT12").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT12", CheckLine, r[3].Value.ToString());
                            }
                        }
                        else
                        {
                            oDS_PH_PY109B.InsertRecord((oDS_PH_PY109B.Size));
                            oDS_PH_PY109B.Offset = oDS_PH_PY109B.Size - 1;
                            oDS_PH_PY109B.SetValue("U_MSTCOD", oDS_PH_PY109B.Size - 1, MSTCOD);
                            oDS_PH_PY109B.SetValue("U_MSTNAM", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
                            oDS_PH_PY109B.SetValue("U_DPTCOD", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(1).Value));
                            oDS_PH_PY109B.SetValue("U_DPTNAM", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(2).Value));

                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT01").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT01", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT02").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT02", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT03").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT03", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT04").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT04", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT05").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT05", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT06").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT06", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT07").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT07", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT08").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT08", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT09").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT09", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT10").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT10", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT11").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT11", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                            if (t[3].Value.ToString().Trim() == oMat1.Columns.Item("AMT12").TitleObject.Caption)
                            {
                                oDS_PH_PY109B.SetValue("U_AMT12", oDS_PH_PY109B.Size - 1, r[3].Value.ToString());
                            }
                        }

                        if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50)
                        {
                            ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                            oProValue = oProValue + 1;
                            ProgressBar01.Value = oProValue;
                        }
                        tRow = tRow + 1;
                    }
                }

                // 수당, 공제 항목 이외의 값은 0으로 처리
                for (i = 0; i <= oDS_PH_PY109B.Size - 1; i++)
                {
                    for (j = 9 + oDS_PH_PY109Z.Size; j <= oDS_PH_PY109B.Fields.Count - 1; j++)
                    {
                        oDS_PH_PY109B.SetValue(j, i, Convert.ToString(0));
                    }
                    oDS_PH_PY109B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    // 라인번호 재정의
                }

                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();

                ProgressBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText("엑셀을 불러왔습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                //진행바 초기화
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oForm.Items.Item("CLTCOD").Enabled = false;
                oForm.Items.Item("YM").Enabled = false;
                oForm.Items.Item("JOBTYP").Enabled = false;
                oForm.Items.Item("JOBGBN").Enabled = false;
                oForm.Items.Item("JOBTRG").Enabled = false;

            }
            catch (Exception ex)
            {
                if ((ProgressBar01 != null))
                {
                    ProgressBar01.Stop();
                    ProgressBar01 = null;
                }

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("항목 설정이 되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀파일이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("항목이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 사번", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B열 두번째 행 타이틀은 성명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + "행의 [" + xlsh.Cells[i, 1] + " ] 사번이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY109_Excel_Upload_ERROR" + MSTCOD + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                xlapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

            }
            finally
            {
                xlapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// Sum_Display
        /// </summary>
        private void Sum_Display()
        {
            int i = 0;

            Double Amt01 = 0;
            Double Amt02 = 0;
            Double Amt03 = 0;
            Double Amt04 = 0;
            Double Amt05 = 0;
            Double Amt06 = 0;
            Double Amt07 = 0;
            Double Amt08 = 0;
            Double Amt09 = 0;
            Double Amt10 = 0;
            Double Amt11 = 0;
            Double Amt12 = 0;

            oMat1.FlushToDataSource();
            oMat1.LoadFromDataSource();
            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
            {
                Amt01 = Amt01 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT01", i));
                Amt02 = Amt02 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT02", i));
                Amt03 = Amt03 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT03", i));
                Amt04 = Amt04 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT04", i));
                Amt05 = Amt05 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT05", i));
                Amt06 = Amt06 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT06", i));
                Amt07 = Amt07 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT07", i));
                Amt08 = Amt08 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT08", i));
                Amt09 = Amt09 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT09", i));
                Amt10 = Amt10 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT10", i));
                Amt11 = Amt11 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT11", i));
                Amt12 = Amt12 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT12", i));

            }

            oForm.Items.Item("T1").Specific.Caption = oMat1.Columns.Item("AMT01").TitleObject.Caption;
            oForm.Items.Item("T2").Specific.Caption = oMat1.Columns.Item("AMT02").TitleObject.Caption;
            oForm.Items.Item("T3").Specific.Caption = oMat1.Columns.Item("AMT03").TitleObject.Caption;
            oForm.Items.Item("T4").Specific.Caption = oMat1.Columns.Item("AMT04").TitleObject.Caption;
            oForm.Items.Item("T5").Specific.Caption = oMat1.Columns.Item("AMT05").TitleObject.Caption;
            oForm.Items.Item("T6").Specific.Caption = oMat1.Columns.Item("AMT06").TitleObject.Caption;
            oForm.Items.Item("T7").Specific.Caption = oMat1.Columns.Item("AMT07").TitleObject.Caption;
            oForm.Items.Item("T8").Specific.Caption = oMat1.Columns.Item("AMT08").TitleObject.Caption;
            oForm.Items.Item("T9").Specific.Caption = oMat1.Columns.Item("AMT09").TitleObject.Caption;
            oForm.Items.Item("T10").Specific.Caption = oMat1.Columns.Item("AMT10").TitleObject.Caption;
            oForm.Items.Item("T11").Specific.Caption = oMat1.Columns.Item("AMT11").TitleObject.Caption;
            oForm.Items.Item("T12").Specific.Caption = oMat1.Columns.Item("AMT12").TitleObject.Caption;

            oForm.DataSources.UserDataSources.Item("S_Amt01").Value = Convert.ToString(Amt01);
            oForm.DataSources.UserDataSources.Item("S_Amt02").Value = Convert.ToString(Amt02);
            oForm.DataSources.UserDataSources.Item("S_Amt03").Value = Convert.ToString(Amt03);
            oForm.DataSources.UserDataSources.Item("S_Amt04").Value = Convert.ToString(Amt04);
            oForm.DataSources.UserDataSources.Item("S_Amt05").Value = Convert.ToString(Amt05);
            oForm.DataSources.UserDataSources.Item("S_Amt06").Value = Convert.ToString(Amt06);
            oForm.DataSources.UserDataSources.Item("S_Amt07").Value = Convert.ToString(Amt07);
            oForm.DataSources.UserDataSources.Item("S_Amt08").Value = Convert.ToString(Amt08);
            oForm.DataSources.UserDataSources.Item("S_Amt09").Value = Convert.ToString(Amt09);
            oForm.DataSources.UserDataSources.Item("S_Amt10").Value = Convert.ToString(Amt10);
            oForm.DataSources.UserDataSources.Item("S_Amt11").Value = Convert.ToString(Amt11);
            oForm.DataSources.UserDataSources.Item("S_Amt12").Value = Convert.ToString(Amt12);
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
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//	internal class PH_PY109
//	{
//////********************************************************************************
//////  File           : PH_PY109.cls
//////  Module         : 인사관리 > 급여관리
//////  Desc           : 급상여변동자료등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;
//		public SAPbouiCOM.Matrix oMat2;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY109A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY109B;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY109Z;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		private string sCode;
//		private string sCLTCOD;
//		private string sYM;
//		private string sJOBTYP;
//		private string sJOBGBN;
//		private string sJOBTRG;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY109.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));

//			oFormUniqueID = "PH_PY109_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY109");

//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "Code";

//			oForm.Freeze(true);
//			PH_PY109_CreateItems();
//			PH_PY109_EnableMenus();
//			PH_PY109_SetDocument(oFromDocEntry01);
//			//    Call PH_PY109_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY109_CreateItems()
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

//			///Matrix
//			oDS_PH_PY109A = oForm.DataSources.DBDataSources("@PH_PY109A");
//			////헤더
//			oDS_PH_PY109B = oForm.DataSources.DBDataSources("@PH_PY109B");
//			////라인
//			oDS_PH_PY109Z = oForm.DataSources.DBDataSources("@PH_PY109Z");
//			////라인

//			oMat1 = oForm.Items.Item("Mat1").Specific;
//			oMat2 = oForm.Items.Item("Mat2").Specific;
//			//

//			oForm.DataSources.UserDataSources.Add("S_Amt01", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt01").Specific.DataBind.SetBound(true, "", "S_Amt01");

//			oForm.DataSources.UserDataSources.Add("S_Amt02", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt02").Specific.DataBind.SetBound(true, "", "S_Amt02");

//			oForm.DataSources.UserDataSources.Add("S_Amt03", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt03").Specific.DataBind.SetBound(true, "", "S_Amt03");

//			oForm.DataSources.UserDataSources.Add("S_Amt04", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt04").Specific.DataBind.SetBound(true, "", "S_Amt04");

//			oForm.DataSources.UserDataSources.Add("S_Amt05", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt05").Specific.DataBind.SetBound(true, "", "S_Amt05");

//			oForm.DataSources.UserDataSources.Add("S_Amt06", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt06").Specific.DataBind.SetBound(true, "", "S_Amt06");

//			oForm.DataSources.UserDataSources.Add("S_Amt07", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt07").Specific.DataBind.SetBound(true, "", "S_Amt07");

//			oForm.DataSources.UserDataSources.Add("S_Amt08", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt08").Specific.DataBind.SetBound(true, "", "S_Amt08");

//			oForm.DataSources.UserDataSources.Add("S_Amt09", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt09").Specific.DataBind.SetBound(true, "", "S_Amt09");

//			oForm.DataSources.UserDataSources.Add("S_Amt10", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt10").Specific.DataBind.SetBound(true, "", "S_Amt10");

//			oForm.DataSources.UserDataSources.Add("S_Amt11", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt11").Specific.DataBind.SetBound(true, "", "S_Amt11");

//			oForm.DataSources.UserDataSources.Add("S_Amt12", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("S_Amt12").Specific.DataBind.SetBound(true, "", "S_Amt12");




//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();
//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();


//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------
//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			/// 귀속년월
//			//UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//			//// 지급종류
//			oCombo = oForm.Items.Item("JOBTYP").Specific;
//			oCombo.ValidValues.Add("1", "급여");
//			oCombo.ValidValues.Add("2", "상여");
//			oForm.Items.Item("JOBTYP").DisplayDesc = true;

//			//// 지급구분
//			oCombo = oForm.Items.Item("JOBGBN").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("JOBGBN").DisplayDesc = true;

//			//// 지급대상
//			oCombo = oForm.Items.Item("JOBTRG").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("JOBTRG").DisplayDesc = true;

//			//// 라인-------------------------------------------------------------------------------------------
//			////사번

//			//    '// 부서명
//			//    Set oColumn = oMat1.Columns("DPTNAM")
//			//    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'"
//			//    oRecordSet.DoQuery sQry
//			//    If oRecordSet.RecordCount > 0 Then
//			//        Do Until oRecordSet.EOF
//			//            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
//			//            oRecordSet.MoveNext
//			//        Loop
//			//    End If
//			//    oColumn.DisplayDesc = True

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
//			PH_PY109_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		private void PH_PY109_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY109_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY109_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY109_FormItemEnabled();
//				PH_PY109_AddMatrixRow();
//				PH_PY109_TitleSetting_Matrix();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY109_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY109_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY109_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Conditions oConditions = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("JOBTYP").Enabled = true;
//				oForm.Items.Item("JOBGBN").Enabled = true;
//				oForm.Items.Item("JOBTRG").Enabled = true;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				/// 귀속년월
//				//UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("YM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//				PH_PY109_AddMatrixRow();
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("JOBTYP").Enabled = true;
//				oForm.Items.Item("JOBGBN").Enabled = true;
//				oForm.Items.Item("JOBTRG").Enabled = true;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oForm.Items.Item("YM").Enabled = false;
//				oForm.Items.Item("JOBTYP").Enabled = false;
//				oForm.Items.Item("JOBGBN").Enabled = false;
//				oForm.Items.Item("JOBTRG").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY109_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			string FullName = null;
//			string FindYN = null;


//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;


//			SAPbouiCOM.Conditions oConditions = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (PH_PY109_DataValidCheck(ref (pval.ItemUID)) == false) {
//									BubbleEvent = false;
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//							}
//						}
//						if (pval.ItemUID == "Btn_Set") {
//							if (PH_PY109_DataValidCheck(ref (pval.ItemUID)) == false) {
//								BubbleEvent = false;
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (pval.ActionSuccess == true) {
//								PH_PY109_FormItemEnabled();
//								PH_PY109_TitleSetting_Matrix();
//							}
//						}
//						if (pval.ItemUID == "Btn_Upload") {
//							oMat1.FlushToDataSource();
//							PH_PY109_Excel_Upload();
//							PH_PY109_AddMatrixRow();

//							Sum_Display();

//						}
//						if (pval.ItemUID == "Btn_Set") {
//							if (pval.ActionSuccess == true) {
//								PH_PY109_LoadData_SudangGongje();
//							}
//						}

//						if (pval.ItemUID == "Search") {
//							FindYN = "N";
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							FullName = oForm.Items.Item("FullName").Specific.VALUE;
//							for (i = 1; i <= oMat1.VisualRowCount - 1; i++) {
//								//UPGRADE_WARNING: oMat1.Columns(MSTNAM).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (oMat1.Columns.Item("MSTNAM").Cells.Item(i).Specific.VALUE == FullName) {
//									FindYN = "Y";
//									oMat1.SelectRow(i, true, false);
//								}
//							}
//							if (FindYN == "N") {
//								MDC_Globals.Sbo_Application.SetStatusBarMessage("찾는 사원이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//							}

//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true & pval.ItemUID == "Mat1" & pval.ColUID == "MSTCOD" & pval.CharPressed == 9) {
//						//UPGRADE_WARNING: oMat1.Columns(MSTCOD).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Specific.VALUE)) {
//							oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						} else {
//							//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", Convert.ToString(Convert.ToDouble("'") + oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Specific.VALUE + Convert.ToDouble("'"))) == true) {
//								oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat1":
//						case "Mat2":
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
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {

//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat1":
//							case "Mat2":
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);

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

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "Mat1" & pval.ColUID == "MSTCOD") {
//								//UPGRADE_WARNING: oMat1.Columns(MSTCOD).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (!string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Specific.VALUE)) {
//									//UPGRADE_WARNING: oMat1.Columns(MSTNAM).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oMat1.Columns.Item("MSTNAM").Cells.Item(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Specific.VALUE + "'");
//									//UPGRADE_WARNING: oMat1.Columns(DPTCOD).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oMat1.Columns.Item("DPTCOD").Cells.Item(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Specific.VALUE + "'");
//									//UPGRADE_WARNING: oMat1.Columns(DPTNAM).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oMat1.Columns.Item("DPTNAM").Cells.Item(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData(ref "U_CodeNm", ref "U_Code", ref "[@PS_HR200L]", ref "'" + oMat1.Columns.Item("DPTCOD").Cells.Item(pval.Row).Specific.VALUE + "'", ref " AND Code = '1'");
//								}
//								PH_PY109_AddMatrixRow();
//								oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							}

//							if (pval.ItemUID == "Mat1" & pval.ColUID == "MSTCOD") {
//								PH_PY109_AddMatrixRow();
//								oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

//						PH_PY109_FormItemEnabled();
//						PH_PY109_AddMatrixRow();
//						PH_PY109_TitleSetting_Matrix();

//						Sum_Display();


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
//						//UPGRADE_NOTE: oDS_PH_PY109A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY109A = null;
//						//UPGRADE_NOTE: oDS_PH_PY109B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY109B = null;
//						//UPGRADE_NOTE: oDS_PH_PY109Z 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY109Z = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//						//UPGRADE_NOTE: oMat2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat2 = null;

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
//					if (pval.BeforeAction == true) {

//					} else if (pval.Before_Action == false) {
//						if (pval.ItemUID == "Mat1") {
//							//// 프로젝트 코드
//							if (pval.ColUID == "MSTCOD") {
//								MDC_SetMod.MDC_CF_DBDatasourceReturn(pval, (pval.FormUID), "@PH_PY109B", "U_MSTCOD,U_MSTNAM,U_DPTCOD,U_DPTNAM", "Mat1", (pval.Row));
//								PH_PY109_AddMatrixRow();
//								oMat1.Columns.Item("MSTCOD").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								oMat1.AutoResizeColumns();
//							}
//						}
//					}
//					break;
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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY109A", ref "Code");
//						////접속자 권한에 따른 사업장 보기
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY109_FormItemEnabled();
//						break;


//					//                Call PH_PY109_AddMatrixRow
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY109_FormItemEnabled();
//						Sum_Display();
//						//                Call PH_PY109_AddMatrixRow
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY109_FormItemEnabled();
//						//                Call PH_PY109_AddMatrixRow
//						Sum_Display();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY109_FormItemEnabled();
//						break;

//					case "1293":
//						//// 행삭제
//						//// [MAT1 용]
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY109B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY109B.GetValue("U_MSTCOD", i))) {
//									oDS_PH_PY109B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY109B.Size; i++) {
//								oDS_PH_PY109B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						break;

//					//                Call PH_PY109_AddMatrixRow

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

//			 // ERROR: Not supported in C#: OnErrorStatement


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

//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY109_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			//    '//[Mat1 용]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", oRow - 1)))) {
//					if (oDS_PH_PY109B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY109B.InsertRecord((oRow));
//					}
//					oDS_PH_PY109B.Offset = oRow;
//					oDS_PH_PY109B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY109B.SetValue("U_MSTCOD", oRow, "");
//					oDS_PH_PY109B.SetValue("U_MSTNAM", oRow, "");
//					oDS_PH_PY109B.SetValue("U_DPTCOD", oRow, "");
//					oDS_PH_PY109B.SetValue("U_DPTNAM", oRow, "");
//					oDS_PH_PY109B.SetValue("U_AMT01", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT02", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT03", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT04", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT05", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT06", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT07", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT08", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT09", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT10", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT11", oRow, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT12", oRow, Convert.ToString(0));
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY109B.Offset = oRow - 1;
//					oDS_PH_PY109B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY109B.SetValue("U_MSTCOD", oRow - 1, "");
//					oDS_PH_PY109B.SetValue("U_MSTNAM", oRow - 1, "");
//					oDS_PH_PY109B.SetValue("U_DPTCOD", oRow - 1, "");
//					oDS_PH_PY109B.SetValue("U_DPTNAM", oRow - 1, "");
//					oDS_PH_PY109B.SetValue("U_AMT01", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT02", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT03", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT04", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT05", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT06", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT07", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT08", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT09", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT10", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT11", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY109B.SetValue("U_AMT12", oRow - 1, Convert.ToString(0));
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY109B.Offset = oRow;
//				oDS_PH_PY109B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY109B.SetValue("U_MSTCOD", oRow, "");
//				oDS_PH_PY109B.SetValue("U_MSTNAM", oRow, "");
//				oDS_PH_PY109B.SetValue("U_DPTCOD", oRow, "");
//				oDS_PH_PY109B.SetValue("U_DPTNAM", oRow, "");
//				oDS_PH_PY109B.SetValue("U_AMT01", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT02", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT03", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT04", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT05", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT06", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT07", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT08", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT09", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT10", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT11", oRow, Convert.ToString(0));
//				oDS_PH_PY109B.SetValue("U_AMT12", oRow, Convert.ToString(0));
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY109_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void PH_PY109_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY109'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY109_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY109_DataValidCheck(ref string ItemUID)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////----------------------------------------------------------------------------------
//			////필수 체크
//			////----------------------------------------------------------------------------------
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109A.GetValue("U_YM", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("귀속년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("지급종류는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("JOBTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("JOBGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("지급대상은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("JOBTRG").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}


//			oDS_PH_PY109A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)) + Strings.Right(Strings.Trim(oDS_PH_PY109A.GetValue("U_YM", 0)), 4) + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)) + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)) + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)));
//			oDS_PH_PY109A.SetValue("Name", 0, Strings.Trim(oDS_PH_PY109A.GetValue("COde", 0)));

//			if (ItemUID == "1" & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(Code, Code, [PH_PY109A], ' & Trim$(oDS_PH_PY109A.GetValue(COde, 0)) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" + Strings.Trim(oDS_PH_PY109A.GetValue("COde", 0)) + "'"))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					return functionReturnValue;
//				}
//			} else if (ItemUID == "Btn_Set") {
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(Code, Code, [PH_PY109A], ' & Trim$(oDS_PH_PY109A.GetValue(COde, 0)) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" + Strings.Trim(oDS_PH_PY109A.GetValue("COde", 0)) + "'"))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					return functionReturnValue;
//				}
//			}

//			oMat1.FlushToDataSource();


//			if (ItemUID == "1") {
//				//// Matrix 마지막 행 삭제(DB 저장시)
//				if (oDS_PH_PY109B.Size > 1) {
//					oDS_PH_PY109B.RemoveRecord((oDS_PH_PY109B.Size - 1));
//				} else {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					functionReturnValue = false;
//					return functionReturnValue;
//				}
//			}
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY109_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY109_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY109A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY109A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY109_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY109_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY109_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private object PH_PY109_LoadData_SudangGongje()
//		{
//			object functionReturnValue = null;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze((true));

//			if (oDS_PH_PY109Z.Size > 0) {
//				oDS_PH_PY109Z.Clear();
//			}

//			sCode = Strings.Trim(oDS_PH_PY109A.GetValue("Code", 0));
//			sCLTCOD = Strings.Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0));
//			sYM = Strings.Trim(oDS_PH_PY109A.GetValue("U_YM", 0));
//			sJOBTYP = Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0));
//			sJOBGBN = Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0));
//			sJOBTRG = Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0));

//			//// 수당, 공제 테이블 고정:V, 상여:Y 인 값을 임시테이블에 넣는다
//			sQry = "EXEC PH_PY109 '" + Strings.Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)) + "' , '" + Strings.Trim(oDS_PH_PY109A.GetValue("U_YM", 0)) + "' , '";
//			sQry = sQry + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)) + "' , '" + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)) + "' , '";
//			sQry = sQry + Strings.Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)) + "'";

//			oRecordSet.DoQuery(sQry);

//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			oCombo.Select("" + sCLTCOD + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("YM").Specific.VALUE = sYM;
//			oCombo = oForm.Items.Item("JOBTYP").Specific;
//			oCombo.Select("" + sJOBTYP + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("JOBGBN").Specific;
//			oCombo.Select("" + sJOBGBN + "", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oCombo = oForm.Items.Item("JOBTRG").Specific;
//			oCombo.Select("" + sJOBTRG + "", SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

//			oForm.Freeze((false));
//			return functionReturnValue;
//			PH_PY109_DataLoad_ERROR:

//			oForm.Freeze((false));
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_DataLoad_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private object PH_PY109_LoadData_User()
//		{
//			object functionReturnValue = null;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze((true));
//			//
//			//    sQry = ""
//			//    oDS_PH_PY109_Grid1.Rows.Clear
//			//
//			//    If oDS_PH_PY109B.Size > 1 Then
//			//        oDS_PH_PY109_Grid1.Rows.Add (oDS_PH_PY109B.Size)
//			//        For i = 0 To oDS_PH_PY109B.Size - 1
//			//            oDS_PH_PY109_Grid1.setValue "U_MSTCOD", i, oDS_PH_PY109B.GetValue("U_MSTCOD", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_MSTNAM", i, oDS_PH_PY109B.GetValue("U_MSTNAM", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_DPTCOD", i, oDS_PH_PY109B.GetValue("U_DPTCOD", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_DPTNAM", i, oDS_PH_PY109B.GetValue("U_DPTNAM", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT01", i, oDS_PH_PY109B.GetValue("U_AMT01", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT02", i, oDS_PH_PY109B.GetValue("U_AMT02", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT03", i, oDS_PH_PY109B.GetValue("U_AMT03", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT04", i, oDS_PH_PY109B.GetValue("U_AMT04", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT05", i, oDS_PH_PY109B.GetValue("U_AMT05", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT06", i, oDS_PH_PY109B.GetValue("U_AMT06", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT07", i, oDS_PH_PY109B.GetValue("U_AMT07", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT08", i, oDS_PH_PY109B.GetValue("U_AMT08", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT09", i, oDS_PH_PY109B.GetValue("U_AMT09", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT10", i, oDS_PH_PY109B.GetValue("U_AMT10", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT11", i, oDS_PH_PY109B.GetValue("U_AMT11", i)
//			//            oDS_PH_PY109_Grid1.setValue "U_AMT12", i, oDS_PH_PY109B.GetValue("U_AMT12", i)
//			//        Next
//			//    End If

//			//    oDS_PH_PY109_Grid1.Rows.Add (1)

//			//    oGrid1.AutoResizeColumns

//			oForm.Freeze((false));
//			return functionReturnValue;
//			PH_PY109_DataLoad_ERROR:

//			oForm.Freeze((false));
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_LoadData_User_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY109_TitleSetting_Matrix()
//		{
//			int i = 0;
//			int iCount = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			iCount = 0;
//			for (i = 5; i <= oMat1.Columns.Count - 1; i++) {
//				if (oDS_PH_PY109Z.Size >= 1) {
//					if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109Z.GetValue("U_PDName", 0)))) {
//						oMat1.Columns.Item(1).Editable = false;
//						oMat1.Columns.Item(i).TitleObject.Caption = "";
//						oMat1.Columns.Item(i).Editable = false;
//					} else {
//						oMat1.Columns.Item(1).Editable = true;
//						if (i > 4 & i <= oDS_PH_PY109Z.Size + 4) {
//							oMat1.Columns.Item(i).TitleObject.Caption = Strings.Trim(oDS_PH_PY109Z.GetValue("U_PDName", iCount));
//							oMat1.Columns.Item(i).Editable = true;
//							iCount = iCount + 1;
//						} else {
//							oMat1.Columns.Item(i).TitleObject.Caption = "";
//							oMat1.Columns.Item(i).Editable = false;
//						}
//					}
//				} else {
//					oMat1.Columns.Item(1).Editable = false;
//					oMat1.Columns.Item(i).TitleObject.Caption = "";
//					oMat1.Columns.Item(i).Editable = false;
//				}
//			}


//			oMat1.AutoResizeColumns();

//			oForm.Freeze(false);


//			return;
//			Error_Message:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY109_TitleSetting_Matrix Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}



//		private void PH_PY109_Excel_Upload()
//		{

//			int i = 0;
//			int j = 0;
//			int k = 0;
//			int CheckLine = 0;
//			bool CheckYN = false;
//			string sPrice = null;
//			string sFile = null;
//			string OneRec = null;
//			string sQry = null;
//			short ErrNum = 0;
//			Microsoft.Office.Interop.Excel.Application xl = default(Microsoft.Office.Interop.Excel.Application);
//			Microsoft.Office.Interop.Excel.Workbook xlwb = default(Microsoft.Office.Interop.Excel.Workbook);
//			Microsoft.Office.Interop.Excel.Worksheet xlsh = default(Microsoft.Office.Interop.Excel.Worksheet);

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.Form oForm = null;

//			int TOTCNT = 0;
//			int V_StatusCnt = 0;
//			int oProValue = 0;
//			int tRow = 0;
//			////progbar

//			SAPbobsCOM.Recordset oRecordSet = null;

//			int Amt01 = 0;
//			int Amt02 = 0;
//			int Amt03 = 0;
//			int Amt04 = 0;
//			int Amt05 = 0;
//			int Amt06 = 0;
//			int Amt07 = 0;
//			int Amt08 = 0;
//			int Amt09 = 0;
//			int Amt10 = 0;
//			int Amt11 = 0;
//			int Amt12 = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm = MDC_Globals.Sbo_Application.Forms.ActiveForm;

//			if (oMat1.Columns.Item("MSTCOD").Editable == false) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("항목 설정이 되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				return;
//			}

//			//UPGRADE_WARNING: FileListBoxForm.OpenDialog() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sFile = My.MyProject.Forms.FileListBoxForm.OpenDialog(ref FileListBoxForm, ref "*.xls", ref "파일선택", ref "C:\\");

//			if (string.IsNullOrEmpty(sFile)) {
//				return;
//			} else {
//				if (Strings.Right(Strings.Replace(sFile, Strings.Chr(0), ""), 3) != "xls" & Strings.Right(Strings.Replace(sFile, Strings.Chr(0), ""), 4) != "xlsx") {
//					MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀파일이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return;
//				}
//			}

//			if ((MDC_Globals.oProgBar != null)) {
//				MDC_Globals.oProgBar.Stop();
//				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				MDC_Globals.oProgBar = null;
//			}

//			//엑셀 Object 연결
//			xl = Interaction.CreateObject("excel.application");
//			xlwb = xl.Workbooks.Open(sFile, , true);
//			xlsh = xlwb.Worksheets("급상여변동");


//			if (xlsh.UsedRange.Columns.Count <= 2) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("항목이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 1).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 1).VALUE != "사번") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 사번", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 2).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 2).VALUE != "성명") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B열 두번째 행 타이틀은 성명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//프로그레스 바    ///////////////////////////////////////
//			MDC_Globals.oProgBar = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);

//			//최대값 구하기
//			TOTCNT = xlsh.UsedRange.Rows.Count - 1;

//			V_StatusCnt = System.Math.Round(TOTCNT / 50, 0);
//			oProValue = 1;
//			tRow = 1;
//			///////////////////////////////////////////////////////

//			for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++) {
//				////사번 존재 여부 체크
//				//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + Strings.Trim(xlsh.Cells._Default(i, 1)) + "'") == true) {
//					ErrNum = 1;
//					goto Err_Renamed;
//				} else {
//					CheckYN = false;

//					for (j = 0; j <= oDS_PH_PY109B.Size - 1; j++) {
//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (Strings.Trim(xlsh.Cells._Default(i, 1)) == Strings.Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", j))) {
//							CheckYN = true;
//							CheckLine = j;
//							break; // TODO: might not be correct. Was : Exit For
//						}
//					}

//					////마지막행 제거
//					if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", oDS_PH_PY109B.Size - 1)))) {
//						oDS_PH_PY109B.RemoveRecord((oDS_PH_PY109B.Size - 1));
//					}

//					//// 사원마스터에서 사번에 대한 정보 가져오기
//					sQry = "select U_FullName, U_TeamCode, U_CodeNm ";
//					sQry = sQry + " FROM [@PH_PY001A] T0 INNER JOIN [@Ps_HR200L] T1 ON T0.U_Teamcode = T1.U_Code";
//					//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE T0.Code='" + Strings.Trim(xlsh.Cells._Default(i, 1)) + "'";
//					oRecordSet.DoQuery(sQry);

//					////Mat1에 업로드엑셀 사번이 존재 유무 CheckYN
//					if (CheckYN == true) {
//						oDS_PH_PY109B.SetValue("U_MSTNAM", CheckLine, Strings.Trim(oRecordSet.Fields.Item(0).Value));
//						oDS_PH_PY109B.SetValue("U_DPTCOD", CheckLine, Strings.Trim(oRecordSet.Fields.Item(1).Value));
//						oDS_PH_PY109B.SetValue("U_DPTNAM", CheckLine, Strings.Trim(oRecordSet.Fields.Item(2).Value));

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt01 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT01"));
//						if (Amt01 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT01", CheckLine, Convert.ToString(Amt01));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt02 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT02"));
//						if (Amt02 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT02", CheckLine, Convert.ToString(Amt02));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt03 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT03"));
//						if (Amt03 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT03", CheckLine, Convert.ToString(Amt03));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt04 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT04"));
//						if (Amt04 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT04", CheckLine, Convert.ToString(Amt04));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt05 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT05"));
//						if (Amt05 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT05", CheckLine, Convert.ToString(Amt05));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt06 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT06"));
//						if (Amt06 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT06", CheckLine, Convert.ToString(Amt06));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt07 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT07"));
//						if (Amt07 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT07", CheckLine, Convert.ToString(Amt07));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt08 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT08"));
//						if (Amt08 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT08", CheckLine, Convert.ToString(Amt08));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt09 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT09"));
//						if (Amt09 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT09", CheckLine, Convert.ToString(Amt09));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt10 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT10"));
//						if (Amt10 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT10", CheckLine, Convert.ToString(Amt10));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt11 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT11"));
//						if (Amt11 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT11", CheckLine, Convert.ToString(Amt11));
//						}

//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						Amt12 = xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT12"));
//						if (Amt12 != 0) {
//							oDS_PH_PY109B.SetValue("U_AMT12", CheckLine, Convert.ToString(Amt12));
//						}

//						//                If TitleCheck(oMat1, xlsh, "AMT01") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT01", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT01"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT02") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT02", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT02"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT03") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT03", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT03"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT04") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT04", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT04"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT05") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT05", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT05"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT06") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT06", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT06"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT07") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT07", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT07"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT08") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT08", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT08"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT09") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT09", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT09"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT10") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT10", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT10"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT11") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT11", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT11"))
//						//                End If
//						//                If TitleCheck(oMat1, xlsh, "AMT12") = True Then
//						//                    oDS_PH_PY109B.setValue "U_AMT12", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT12"))
//						//                End If
//					////새로 사번 추가
//					} else {
//						//                If oDS_PH_PY109B.Size <= oMat1.VisualRowCount Then
//						oDS_PH_PY109B.InsertRecord((oDS_PH_PY109B.Size));
//						//                End If
//						oDS_PH_PY109B.Offset = oDS_PH_PY109B.Size - 1;
//						//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						oDS_PH_PY109B.SetValue("U_MSTCOD", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, 1));
//						oDS_PH_PY109B.SetValue("U_MSTNAM", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(0).Value));
//						oDS_PH_PY109B.SetValue("U_DPTCOD", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(1).Value));
//						oDS_PH_PY109B.SetValue("U_DPTNAM", oDS_PH_PY109B.Size - 1, Strings.Trim(oRecordSet.Fields.Item(2).Value));
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT01") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT01", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT01")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT02") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT02", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT02")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT03") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT03", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT03")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT04") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT04", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT04")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT05") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT05", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT05")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT06") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT06", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT06")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT07") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT07", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT07")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT08") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT08", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT08")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT09") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT09", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT09")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT10") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT10", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT10")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT11") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT11", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT11")));
//						}
//						if (TitleCheck(ref oMat1, ref xlsh, ref "AMT12") == true) {
//							//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oDS_PH_PY109B.SetValue("U_AMT12", oDS_PH_PY109B.Size - 1, xlsh.Cells._Default(i, ValueCheck(ref oMat1, ref xlsh, ref "AMT12")));
//						}
//					}

//					if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50) {
//						MDC_Globals.oProgBar.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
//						oProValue = oProValue + 1;
//						MDC_Globals.oProgBar.Value = oProValue;
//					}
//					tRow = tRow + 1;
//				}
//			}

//			////수당, 공제 항목 이외의 값은 0으로 처리
//			for (i = 0; i <= oDS_PH_PY109B.Size - 1; i++) {
//				for (j = 9 + oDS_PH_PY109Z.Size; j <= oDS_PH_PY109B.Fields.Count - 1; j++) {
//					oDS_PH_PY109B.SetValue(j, i, Convert.ToString(0));
//				}
//				oDS_PH_PY109B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				////라인번호 재정의
//			}

//			oMat1.LoadFromDataSource();
//			oMat1.AutoResizeColumns();


//			MDC_Globals.oProgBar.Stop();
//			MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀을 불러왔습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

//			//액셀개체 닫음
//			xlwb.Close();
//			//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlwb = null;
//			//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xl = null;
//			//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlsh = null;
//			//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			MDC_Globals.oProgBar = null;
//			//진행바 초기화
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
//			oForm.Items.Item("CLTCOD").Enabled = false;
//			oForm.Items.Item("YM").Enabled = false;
//			oForm.Items.Item("JOBTYP").Enabled = false;
//			oForm.Items.Item("JOBGBN").Enabled = false;
//			oForm.Items.Item("JOBTRG").Enabled = false;

//			return;
//			Err_Renamed:

//			if ((MDC_Globals.oProgBar != null)) {
//				MDC_Globals.oProgBar.Stop();
//				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				MDC_Globals.oProgBar = null;
//			}
//			if (ErrNum == 1) {
//				//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.Sbo_Application.StatusBar.SetText(i + "행의 [" + xlsh.Cells._Default(i, 1) + " ] 사번이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText(Err().Description);
//			}
//			//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			MDC_Globals.oProgBar = null;
//			xlwb.Close();
//			//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlwb = null;
//			//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xl = null;
//			//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlsh = null;

//		}

//		private int ValueCheck(ref SAPbouiCOM.Matrix oMat1, ref Microsoft.Office.Interop.Excel.Worksheet xlsh, ref string Field)
//		{
//			int functionReturnValue = 0;
//			int i = 0;
//			bool check = false;

//			check = false;
//			for (i = 3; i <= xlsh.UsedRange.Columns.Count; i++) {
//				//UPGRADE_WARNING: xlsh.Cells(1, i).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(xlsh.Cells._Default(1, i).VALUE)) {
//					if (check == false) {
//						//UPGRADE_WARNING: xlsh.Cells(1, i).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (xlsh.Cells._Default(1, i).VALUE == oMat1.Columns.Item(Field).Title) {
//							functionReturnValue = i;
//							check = true;
//							if (functionReturnValue <= 2) {
//								functionReturnValue = 50;
//								//16384
//							}
//							break; // TODO: might not be correct. Was : Exit For
//						}
//					} else {
//						MDC_Globals.Sbo_Application.StatusBar.SetText("항목에 중복된 이름이 있습니다.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					}
//				}
//			}

//			////해당 항목이 없으면 마지막 컬럼값(빈값)의 0 값을 가져오게 설정
//			if (functionReturnValue <= 2) {
//				functionReturnValue = 50;
//				//16384
//			}
//			return functionReturnValue;
//		}


//		private bool TitleCheck(ref SAPbouiCOM.Matrix oMat1, ref Microsoft.Office.Interop.Excel.Worksheet xlsh, ref string Field)
//		{
//			bool functionReturnValue = false;
//			int i = 0;
//			bool check = false;

//			check = false;
//			for (i = 3; i <= xlsh.UsedRange.Columns.Count; i++) {
//				//UPGRADE_WARNING: xlsh.Cells(1, i).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(xlsh.Cells._Default(1, i).VALUE)) {
//					if (check == false) {
//						//UPGRADE_WARNING: xlsh.Cells(1, i).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (xlsh.Cells._Default(1, i).VALUE == oMat1.Columns.Item(Field).Title) {
//							functionReturnValue = true;
//							check = true;
//						}
//					} else {
//						//Sbo_Application.StatusBar.SetText "항목에 중복된 이름이 있습니다.", bmt_Long, smt_Error
//					}
//				}
//			}
//			return functionReturnValue;
//		}


//		private int Sum_Display()
//		{
//			int i = 0;


//			int Amt01 = 0;
//			int Amt02 = 0;
//			int Amt03 = 0;
//			int Amt04 = 0;
//			int Amt05 = 0;
//			int Amt06 = 0;
//			int Amt07 = 0;
//			int Amt08 = 0;
//			int Amt09 = 0;
//			int Amt10 = 0;
//			int Amt11 = 0;
//			int Amt12 = 0;

//			Amt01 = 0;
//			Amt02 = 0;
//			Amt03 = 0;
//			Amt04 = 0;
//			Amt05 = 0;
//			Amt06 = 0;
//			Amt07 = 0;
//			Amt08 = 0;
//			Amt09 = 0;
//			Amt10 = 0;
//			Amt11 = 0;
//			Amt12 = 0;

//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();
//			for (i = 0; i <= oMat1.VisualRowCount - 1; i++) {
//				Amt01 = Amt01 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT01", i));
//				Amt02 = Amt02 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT02", i));
//				Amt03 = Amt03 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT03", i));
//				Amt04 = Amt04 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT04", i));
//				Amt05 = Amt05 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT05", i));
//				Amt06 = Amt06 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT06", i));
//				Amt07 = Amt07 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT07", i));
//				Amt08 = Amt08 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT08", i));
//				Amt09 = Amt09 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT09", i));
//				Amt10 = Amt10 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT10", i));
//				Amt11 = Amt11 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT11", i));
//				Amt12 = Amt12 + Convert.ToDouble(oDS_PH_PY109B.GetValue("U_AMT12", i));

//			}

//			//UPGRADE_WARNING: oForm.Items(T1).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T1").Specific.Caption = oMat1.Columns.Item("AMT01").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T2).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T2").Specific.Caption = oMat1.Columns.Item("AMT02").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T3).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T3").Specific.Caption = oMat1.Columns.Item("AMT03").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T4).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T4").Specific.Caption = oMat1.Columns.Item("AMT04").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T5).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T5").Specific.Caption = oMat1.Columns.Item("AMT05").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T6).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T6").Specific.Caption = oMat1.Columns.Item("AMT06").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T7).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T7").Specific.Caption = oMat1.Columns.Item("AMT07").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T8).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T8").Specific.Caption = oMat1.Columns.Item("AMT08").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T9).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T9").Specific.Caption = oMat1.Columns.Item("AMT09").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T10).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T10").Specific.Caption = oMat1.Columns.Item("AMT10").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T11).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T11").Specific.Caption = oMat1.Columns.Item("AMT11").TitleObject.Caption;
//			//UPGRADE_WARNING: oForm.Items(T12).Specific.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("T12").Specific.Caption = oMat1.Columns.Item("AMT12").TitleObject.Caption;

//			oForm.DataSources.UserDataSources.Item("S_Amt01").Value = Convert.ToString(Amt01);
//			oForm.DataSources.UserDataSources.Item("S_Amt02").Value = Convert.ToString(Amt02);
//			oForm.DataSources.UserDataSources.Item("S_Amt03").Value = Convert.ToString(Amt03);
//			oForm.DataSources.UserDataSources.Item("S_Amt04").Value = Convert.ToString(Amt04);
//			oForm.DataSources.UserDataSources.Item("S_Amt05").Value = Convert.ToString(Amt05);
//			oForm.DataSources.UserDataSources.Item("S_Amt06").Value = Convert.ToString(Amt06);
//			oForm.DataSources.UserDataSources.Item("S_Amt07").Value = Convert.ToString(Amt07);
//			oForm.DataSources.UserDataSources.Item("S_Amt08").Value = Convert.ToString(Amt08);
//			oForm.DataSources.UserDataSources.Item("S_Amt09").Value = Convert.ToString(Amt09);
//			oForm.DataSources.UserDataSources.Item("S_Amt10").Value = Convert.ToString(Amt10);
//			oForm.DataSources.UserDataSources.Item("S_Amt11").Value = Convert.ToString(Amt11);
//			oForm.DataSources.UserDataSources.Item("S_Amt12").Value = Convert.ToString(Amt12);


//		}
//	}
//}
