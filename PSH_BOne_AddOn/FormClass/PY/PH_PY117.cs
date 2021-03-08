using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
//using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여마감작업
    /// </summary>
    internal class PH_PY117 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY117;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY117.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY117_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY117");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY117_CreateItems();
                PH_PY117_FormItemEnabled();
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
        private void PH_PY117_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY117");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY117");
                oDS_PH_PY117 = oForm.DataSources.DataTables.Item("PH_PY117");

                //그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("마감", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("총지급액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add("실지급액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //귀속년월
                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_DATE, 6);
                oForm.Items.Item("YM").Specific.String = DateTime.Now.ToString("yyyyMM");

                //지급종류
                oForm.DataSources.UserDataSources.Add("JOBTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("JOBTYP").Specific.DataBind.SetBound(true, "", "JOBTYP");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                //지급구분
                oForm.DataSources.UserDataSources.Add("JOBGBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("JOBGBN").Specific.DataBind.SetBound(true, "", "JOBGBN");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "N");
                oForm.Items.Item("JOBGBN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                //지급대상
                oForm.DataSources.UserDataSources.Add("PAYSEL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("PAYSEL").Specific.DataBind.SetBound(true, "", "PAYSEL");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYSEL").Specific, "N");
                oForm.Items.Item("PAYSEL").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PAYSEL").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("PAYSEL").DisplayDesc = true;

                //사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

                //마감
                oForm.DataSources.UserDataSources.Add("Close", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Close").Specific.DataBind.SetBound(true, "", "Close");
                oForm.Items.Item("Close").Specific.ValOn = "Y";
                oForm.Items.Item("Close").Specific.ValOff = "N";
                oForm.Items.Item("Close").Specific.Checked = false;

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY117_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY117_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY117_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY117_DataFind
        /// </summary>
        private void PH_PY117_DataFind()
        {
            short ErrNum = 0;
            int iRow;
            string sQry;
            string CLTCOD;
            string YM;
            string JOBTYP;
            string JOBGBN;
            string PAYSEL;
            string MSTCOD;
            string TeamCode;
            string RspCode;
            string EndCheck;
            
            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                JOBTYP = oForm.Items.Item("JOBTYP").Specific.Value.ToString().Trim();
                JOBGBN = oForm.Items.Item("JOBGBN").Specific.Value.ToString().Trim();
                PAYSEL = oForm.Items.Item("PAYSEL").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                //Close = oForm.Items.Item("Close").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(MSTCOD))
                { 
                    MSTCOD = "%"; 
                }

                EndCheck = oForm.DataSources.UserDataSources.Item("Close").Value.ToString().Trim();  //Check Box
                if (string.IsNullOrEmpty(EndCheck))
                { 
                    EndCheck = "N"; 
                }

                if (string.IsNullOrEmpty(YM))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "Exec PH_PY117 '" + CLTCOD + "','" + YM + "','" + JOBTYP + "','" + JOBGBN + "','" + PAYSEL + "','" + MSTCOD + "','" + TeamCode + "','" + RspCode + "','" + EndCheck + "'";
                oDS_PH_PY117.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY117_TitleSetting(iRow);
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("귀속년월을 선택하세요, 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY117_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY117_DataChange
        /// </summary>
        private void PH_PY117_DataChange()
        {
            string sQry;
            int i;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY117.GetValue("MAGAM", i) == "Y")
                        {
                            sQry = "  UPDATE [@PH_PY112A] SET U_ENDCHK = 'Y' WHERE U_MSTCOD = '" + oDS_PH_PY117.GetValue("U_MSTCOD", i) + "'";
                            sQry += " AND U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                            sQry += " AND U_YM = '" + oForm.Items.Item("YM").Specific.Value.Trim() + "'";
                            sQry += " AND U_JOBTYP = '" + oForm.Items.Item("JOBTYP").Specific.Value.Trim() + "'";
                            sQry += " AND U_JOBGBN = '" + oForm.Items.Item("JOBGBN").Specific.Value.Trim() + "'";
                            sQry += " AND (U_JOBTRG = '" + oForm.Items.Item("PAYSEL").Specific.Value.Trim() + "'";
                            sQry += " OR (U_JOBTRG <> '" + oForm.Items.Item("PAYSEL").Specific.Value.Trim() + "' AND U_JOBTRG LIKE '" + oForm.Items.Item("PAYSEL").Specific.Value.Trim() + "'))";

                            oRecordSet.DoQuery(sQry);
                        }
                    }
                    PSH_Globals.SBO_Application.SetStatusBarMessage("마감처리가 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY117_DataChange_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY117_TitleSetting(int iRow)
        {
            int i;
            string[] COLNAM = new string[80];

            try
            {
                string EndCheck = string.Empty;

                //그리드 콤보박스
                COLNAM[0] = "마감";
                COLNAM[1] = "부서";
                COLNAM[2] = "담당";
                COLNAM[3] = "사번";
                COLNAM[4] = "성명";
                COLNAM[5] = "총지급액";
                COLNAM[6] = "총공제액";
                COLNAM[7] = "실지급액";
                for (i = 1; i <= 36; i++)
                {
                    COLNAM[i + 7] = "지급항목" + i;
                }
                for (i = 1; i <= 36; i++)
                {
                    COLNAM[i + 43] = "공제항목" + i;
                }

                EndCheck = oForm.DataSources.UserDataSources.Item("Close").Value.ToString().Trim();  //Check Box
                if (string.IsNullOrEmpty(EndCheck))
                { EndCheck = "N"; }

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i < COLNAM.Length)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    if (i == 0)
                    {
                        if (EndCheck == "N")
                        {
                            oGrid1.Columns.Item(i).Editable = true;
                        }
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY117_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        PH_PY117_DataFind();
                    }
                    if (pVal.ItemUID == "Btn_All")
                    {
                        oForm.Freeze(true);
                        for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                        {
                            oDS_PH_PY117.SetValue("MAGAM", i, "Y");
                        }
                        oForm.Freeze(false);
                    }
                    if (pVal.ItemUID == "Btn_Rev")
                    {
                        oForm.Freeze(true);
                        for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                        {
                            oDS_PH_PY117.SetValue("MAGAM", i, "N");
                        }
                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("마감처리를 진행 하시겠습니까?", 2, "Yes", "No") == 2)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PH_PY117_DataChange();
                        PH_PY117_DataFind();
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
                            // 기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }
                            sQry = "  SELECT '%' AS [Code], '전체' AS [Name], -1 AS [Seq] UNION ALL ";
                            sQry += " SELECT U_Code AS [Code], U_CodeNm AS [Name],  U_Seq AS [Seq] FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry += " ORDER BY Seq";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");
                            oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("TeamCode").DisplayDesc = true;
                        }
                        //부서가 바뀌면 담당 재설정
                        if (pVal.ItemUID == "TeamCode")
                        {
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }
                            //현재 부서로 다시 Qry
                            sQry = "  SELECT '%' AS [Code], '전체' AS [Name], -1 AS [Seq] UNION ALL  ";
                            sQry += " SELECT U_Code AS [Code], U_CodeNm AS [Name], U_Seq AS [Seq]";
                            sQry += " FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                            sQry += " ORDER BY Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");
                            oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("RspCode").DisplayDesc = true;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY117);
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
