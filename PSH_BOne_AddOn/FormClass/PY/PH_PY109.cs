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
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
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
                PH_PY109_SetDocument(oFormDocEntry01);
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
                oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");

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
        private void PH_PY109_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry01)))
                {
                    PH_PY109_FormItemEnabled();
                    PH_PY109_AddMatrixRow();
                    PH_PY109_TitleSetting_Matrix();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY109_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
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
                    oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");

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
                //    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                        FullName = oForm.Items.Item("FullName").Specific.Value;
                        for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                        {
                            if (oMat1.Columns.Item("MSTNAM").Cells.Item(i).Specific.Value.Trim() == FullName)
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
                    if (string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.Trim()))
                    {
                        oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                        BubbleEvent = false;
                    }
                    else
                    {
                        if (dataHelpClass.Value_ChkYn("[@PH_PY001A]", "Code", Convert.ToString(Convert.ToDouble("'") + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + Convert.ToDouble("'")),"") == true)
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
                            if (!string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.Trim()))
                            {
                                oMat1.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.Trim() + "'","");
                                oMat1.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.Trim() + "'","");
                                oMat1.Columns.Item("DPTNAM").Cells.Item(pVal.Row).Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oMat1.Columns.Item("DPTCOD").Cells.Item(pVal.Row).Specific.Value + "'", " AND Code = '1'");
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
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.Value.ToString().Trim();
                JOBGBN = oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim();
                JOBTRG = oForm.Items.Item("JOBTRG").Specific.Value.ToString().Trim();

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
                oForm.Items.Item("YM").Specific.Value = sYM;
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

