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
    /// 전근무지등록
    /// </summary>
    internal class PH_PY401 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY401;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY401.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY401_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY401");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY401_CreateItems();
                PH_PY401_FormItemEnabled();
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
        private void PH_PY401_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY401");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY401");
                oDS_PH_PY401 = oForm.DataSources.DataTables.Item("PH_PY401");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY401").Columns.Add("직급", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                
                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
                oForm.DataSources.UserDataSources.Item("Year").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
               
                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 사업자등록번호1,2
                oForm.DataSources.UserDataSources.Add("entno1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("entno1").Specific.DataBind.SetBound(true, "", "entno1");
                oForm.DataSources.UserDataSources.Add("entno2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("entno2").Specific.DataBind.SetBound(true, "", "entno2");

                // 근무처명1,2
                oForm.DataSources.UserDataSources.Add("servcomp1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("servcomp1").Specific.DataBind.SetBound(true, "", "servcomp1");
                oForm.DataSources.UserDataSources.Add("servcomp2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("servcomp2").Specific.DataBind.SetBound(true, "", "servcomp2");

                // 시작근무일1,2
                oForm.DataSources.UserDataSources.Add("symd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("symd1").Specific.DataBind.SetBound(true, "", "symd1");
                oForm.DataSources.UserDataSources.Add("symd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("symd2").Specific.DataBind.SetBound(true, "", "symd2");

                // 종료근무일1,2
                oForm.DataSources.UserDataSources.Add("eymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("eymd1").Specific.DataBind.SetBound(true, "", "eymd1");
                oForm.DataSources.UserDataSources.Add("eymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("eymd2").Specific.DataBind.SetBound(true, "", "eymd2");

                // 급여총액1,2
                oForm.DataSources.UserDataSources.Add("payrtot1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("payrtot1").Specific.DataBind.SetBound(true, "", "payrtot1");
                oForm.DataSources.UserDataSources.Add("payrtot2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("payrtot2").Specific.DataBind.SetBound(true, "", "payrtot2");

                // 상여총액1,2
                oForm.DataSources.UserDataSources.Add("bnstot1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bnstot1").Specific.DataBind.SetBound(true, "", "bnstot1");
                oForm.DataSources.UserDataSources.Add("bnstot2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bnstot2").Specific.DataBind.SetBound(true, "", "bnstot2");

                // 국외근로1,2
                oForm.DataSources.UserDataSources.Add("fwork1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fwork1").Specific.DataBind.SetBound(true, "", "fwork1");
                oForm.DataSources.UserDataSources.Add("fwork2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fwork2").Specific.DataBind.SetBound(true, "", "fwork2");

                // 야간근로수당등1,2
                oForm.DataSources.UserDataSources.Add("ndtalw1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ndtalw1").Specific.DataBind.SetBound(true, "", "ndtalw1");
                oForm.DataSources.UserDataSources.Add("ndtalw2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ndtalw2").Specific.DataBind.SetBound(true, "", "ndtalw2");

                // 기타비과세1,2
                oForm.DataSources.UserDataSources.Add("etcntax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("etcntax1").Specific.DataBind.SetBound(true, "", "etcntax1");
                oForm.DataSources.UserDataSources.Add("etcntax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("etcntax2").Specific.DataBind.SetBound(true, "", "etcntax2");

                // 중식수당1,2
                oForm.DataSources.UserDataSources.Add("lnchalw1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("lnchalw1").Specific.DataBind.SetBound(true, "", "lnchalw1");
                oForm.DataSources.UserDataSources.Add("lnchalw2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("lnchalw2").Specific.DataBind.SetBound(true, "", "lnchalw2");

                // 외국납부세액1,2
                oForm.DataSources.UserDataSources.Add("ftaxamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ftaxamt1").Specific.DataBind.SetBound(true, "", "ftaxamt1");
                oForm.DataSources.UserDataSources.Add("ftaxamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ftaxamt2").Specific.DataBind.SetBound(true, "", "ftaxamt2");

                // 저축세액공제1,2
                oForm.DataSources.UserDataSources.Add("savtaxddc1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("savtaxddc1").Specific.DataBind.SetBound(true, "", "savtaxddc1");
                oForm.DataSources.UserDataSources.Add("savtaxddc2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("savtaxddc2").Specific.DataBind.SetBound(true, "", "savtaxddc2");

                // 소득세1,2
                oForm.DataSources.UserDataSources.Add("incmtax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("incmtax1").Specific.DataBind.SetBound(true, "", "incmtax1");
                oForm.DataSources.UserDataSources.Add("indmtax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("indmtax2").Specific.DataBind.SetBound(true, "", "indmtax2");

                // 농어촌특별세1,2
                oForm.DataSources.UserDataSources.Add("fvsptax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fvsptax1").Specific.DataBind.SetBound(true, "", "fvsptax1");
                oForm.DataSources.UserDataSources.Add("fvsptax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("fvsptax2").Specific.DataBind.SetBound(true, "", "fvsptax2");

                // 주민세1,2
                oForm.DataSources.UserDataSources.Add("residtax1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("residtax1").Specific.DataBind.SetBound(true, "", "residtax1");
                oForm.DataSources.UserDataSources.Add("residtax2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("residtax2").Specific.DataBind.SetBound(true, "", "residtax2");

                // 건강보험1,2
                oForm.DataSources.UserDataSources.Add("medcinsr1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcinsr1").Specific.DataBind.SetBound(true, "", "medcinsr1");
                oForm.DataSources.UserDataSources.Add("medcinsr2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcinsr2").Specific.DataBind.SetBound(true, "", "medcinsr2");

                // 고용보험1,2
                oForm.DataSources.UserDataSources.Add("asopinsr1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("asopinsr1").Specific.DataBind.SetBound(true, "", "asopinsr1");
                oForm.DataSources.UserDataSources.Add("asopinsr2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("asopinsr2").Specific.DataBind.SetBound(true, "", "asopinsr2");

                // 연금보험료공제1,2
                oForm.DataSources.UserDataSources.Add("annuboamt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("annuboamt1").Specific.DataBind.SetBound(true, "", "annuboamt1");
                oForm.DataSources.UserDataSources.Add("annuboamt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("annuboamt2").Specific.DataBind.SetBound(true, "", "annuboamt2");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY401_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                //
                oForm.Items.Item("MSTCOD").Specific.VALUE = "";
                oForm.Items.Item("FullName").Specific.VALUE = "";
                oForm.Items.Item("TeamName").Specific.VALUE = "";
                oForm.Items.Item("RspName").Specific.VALUE = "";
                oForm.Items.Item("ClsName").Specific.VALUE = "";

                oForm.DataSources.UserDataSources.Item("entno1").Value = "";
                oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
                oForm.DataSources.UserDataSources.Item("symd1").Value = "";
                oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

                oForm.DataSources.UserDataSources.Item("payrtot1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("bnstot1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("fwork1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("etcntax1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("incmtax1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("residtax1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = Convert.ToString(0);

                oForm.DataSources.UserDataSources.Item("entno2").Value = "";
                oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
                oForm.DataSources.UserDataSources.Item("symd2").Value = "";
                oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

                oForm.DataSources.UserDataSources.Item("payrtot2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("bnstot2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("fwork2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("etcntax2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("indmtax2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("residtax2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = Convert.ToString(0);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY401_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY401);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
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
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY401_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY401_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY401_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY401_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
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
            string sQry = string.Empty;
            string yyyy, Result = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY401_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        yyyy = oForm.Items.Item("Year").Specific.VALUE;
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY401_SAVE();
                        }
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        yyyy = oForm.Items.Item("Year").Specific.VALUE;
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY401_Delete();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
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
                            case "MSTCOD":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();

                                sQry = "Select Code,";
                                sQry = sQry + " FullName = U_FullName,";
                                sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '1'";
                                sQry = sQry + " And U_Code = U_TeamCode),''),";
                                sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '2'";
                                sQry = sQry + " And U_Code = U_RspCode),''),";
                                sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '9'";
                                sQry = sQry + " And U_Code  = U_ClsCode";
                                sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
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

                            Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                            Param02 = oDS_PH_PY401.Columns.Item("Year").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY401.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY401_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                            oRecordSet.DoQuery(sQry);

                            if ((oRecordSet.RecordCount == 0))
                            {
                                oForm.Items.Item("MSTCOD").Specific.VALUE = oDS_PH_PY401.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                                oForm.Items.Item("FullName").Specific.VALUE = oDS_PH_PY401.Columns.Item("FullName").Cells.Item(pVal.Row).Value;

                                oForm.DataSources.UserDataSources.Item("entno1").Value = "";
                                oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
                                oForm.DataSources.UserDataSources.Item("symd1").Value = "";
                                oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

                                oForm.DataSources.UserDataSources.Item("payrtot1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("bnstot1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("fwork1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("etcntax1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("incmtax1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("residtax1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = Convert.ToString(0);

                                oForm.DataSources.UserDataSources.Item("entno2").Value = "";
                                oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
                                oForm.DataSources.UserDataSources.Item("symd2").Value = "";
                                oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

                                oForm.DataSources.UserDataSources.Item("payrtot2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("bnstot2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("fwork2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("etcntax2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("indmtax2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("residtax2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = Convert.ToString(0);

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {
                                oForm.Items.Item("MSTCOD").Specific.VALUE = oRecordSet.Fields.Item("MSTCOD").Value;
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;

                                oForm.DataSources.UserDataSources.Item("entno1").Value = oRecordSet.Fields.Item("entno1").Value;
                                oForm.DataSources.UserDataSources.Item("servcomp1").Value = oRecordSet.Fields.Item("servcomp1").Value;
                                oForm.DataSources.UserDataSources.Item("symd1").Value = Convert.ToDateTime(oRecordSet.Fields.Item("symd1").Value.ToString().Trim()).ToString("yyyyMMdd");
                                oForm.DataSources.UserDataSources.Item("eymd1").Value = Convert.ToDateTime(oRecordSet.Fields.Item("eymd1").Value.ToString().Trim()).ToString("yyyyMMdd");

                                oForm.DataSources.UserDataSources.Item("payrtot1").Value = oRecordSet.Fields.Item("payrtot1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("bnstot1").Value = oRecordSet.Fields.Item("bnstot1").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("fwork1").Value = oRecordSet.Fields.Item("fwork1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ndtalw1").Value = oRecordSet.Fields.Item("ndtalw1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("etcntax1").Value = oRecordSet.Fields.Item("etcntax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("lnchalw1").Value = oRecordSet.Fields.Item("lnchalw1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = oRecordSet.Fields.Item("ftaxamt1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = oRecordSet.Fields.Item("savtaxddc1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("incmtax1").Value = oRecordSet.Fields.Item("incmtax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("fvsptax1").Value = oRecordSet.Fields.Item("fvsptax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("residtax1").Value = oRecordSet.Fields.Item("residtax1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("medcinsr1").Value = oRecordSet.Fields.Item("medcinsr1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("asopinsr1").Value = oRecordSet.Fields.Item("asopinsr1").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("annuboamt1").Value = oRecordSet.Fields.Item("annuboamt1").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("entno2").Value = oRecordSet.Fields.Item("entno2").Value;
                                oForm.DataSources.UserDataSources.Item("servcomp2").Value = oRecordSet.Fields.Item("servcomp2").Value;
                                oForm.DataSources.UserDataSources.Item("symd2").Value = Convert.ToDateTime(oRecordSet.Fields.Item("symd2").Value.ToString().Trim()).ToString("yyyyMMdd");
                                oForm.DataSources.UserDataSources.Item("eymd2").Value = Convert.ToDateTime(oRecordSet.Fields.Item("eymd2").Value.ToString().Trim()).ToString("yyyyMMdd");

                                oForm.DataSources.UserDataSources.Item("payrtot2").Value = oRecordSet.Fields.Item("payrtot2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("bnstot2").Value = oRecordSet.Fields.Item("bnstot2").Value.ToString();

                                oForm.DataSources.UserDataSources.Item("fwork2").Value = oRecordSet.Fields.Item("fwork2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ndtalw2").Value = oRecordSet.Fields.Item("ndtalw2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("etcntax2").Value = oRecordSet.Fields.Item("etcntax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("lnchalw2").Value = oRecordSet.Fields.Item("lnchalw2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = oRecordSet.Fields.Item("ftaxamt2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = oRecordSet.Fields.Item("savtaxddc2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("indmtax2").Value = oRecordSet.Fields.Item("indmtax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("fvsptax2").Value = oRecordSet.Fields.Item("fvsptax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("residtax2").Value = oRecordSet.Fields.Item("residtax2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("medcinsr2").Value = oRecordSet.Fields.Item("medcinsr2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("asopinsr2").Value = oRecordSet.Fields.Item("asopinsr2").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("annuboamt2").Value = oRecordSet.Fields.Item("annuboamt2").Value.ToString();

                                // 부서
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
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
        /// PH_PY401_DataFind
        /// </summary>
        private void PH_PY401_DataFind()
        {
            int iRow = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string Year = string.Empty;

            
            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
            Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();


            if (string.IsNullOrEmpty(Strings.Trim(Year)))
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("년도가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //functionReturnValue = false;
                return; //functionReturnValue;
            }

            try
            {
                oForm.Freeze(true);

                PH_PY401_FormItemEnabled();

                sQry = "Exec PH_PY401_01 '" + CLTCOD + "','" + Year + "'";
                oDS_PH_PY401.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY401_TitleSetting(ref iRow);

                oForm.EnableMenu("1282", true);     //문서추가 활성
            }
            catch (Exception ex)
            {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY401_SAVE
        /// </summary>
        private void PH_PY401_SAVE()
        {
            // 데이타 저장
            short ErrNum = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string CLTCOD, MSTCOD, FullName, YEAR = string.Empty;
            string entno1, servcomp1, symd1, eymd1, entno2, servcomp2, symd2, eymd2 = string.Empty;

            double payrtot1, bnstot1, fwork1, ndtalw1, etcntax1, lnchalw1, ftaxamt1, savtaxddc1, incmtax1, fvsptax1, residtax1, medcinsr1, asopinsr1, annuboamt1 = 0;
            double payrtot2, bnstot2, fwork2, ndtalw2, etcntax2, lnchalw2, ftaxamt2, savtaxddc2, indmtax2, fvsptax2, residtax2, medcinsr2, asopinsr2, annuboamt2 = 0;

            try
            {
                oForm.Freeze(true);

                entno1 = oForm.Items.Item("entno1").Specific.VALUE.ToString().Trim();
                servcomp1 = oForm.Items.Item("servcomp1").Specific.VALUE.ToString().Trim();
                symd1 = oForm.Items.Item("symd1").Specific.VALUE.ToString().Trim();
                eymd1 = oForm.Items.Item("eymd1").Specific.VALUE.ToString().Trim();
                payrtot1 = Convert.ToDouble(oForm.Items.Item("payrtot1").Specific.VALUE);
                bnstot1 = Convert.ToDouble(oForm.Items.Item("bnstot1").Specific.VALUE);
                fwork1 = Convert.ToDouble(oForm.Items.Item("fwork1").Specific.VALUE);
                ndtalw1 = Convert.ToDouble(oForm.Items.Item("ndtalw1").Specific.VALUE);
                etcntax1 = Convert.ToDouble(oForm.Items.Item("etcntax1").Specific.VALUE);
                lnchalw1 = Convert.ToDouble(oForm.Items.Item("lnchalw1").Specific.VALUE);
                ftaxamt1 = Convert.ToDouble(oForm.Items.Item("ftaxamt1").Specific.VALUE);
                savtaxddc1 = Convert.ToDouble(oForm.Items.Item("savtaxddc1").Specific.VALUE);
                incmtax1 = Convert.ToDouble(oForm.Items.Item("incmtax1").Specific.VALUE);
                fvsptax1 = Convert.ToDouble(oForm.Items.Item("fvsptax1").Specific.VALUE);
                residtax1 = Convert.ToDouble(oForm.Items.Item("residtax1").Specific.VALUE);
                medcinsr1 = Convert.ToDouble(oForm.Items.Item("medcinsr1").Specific.VALUE);
                asopinsr1 = Convert.ToDouble(oForm.Items.Item("asopinsr1").Specific.VALUE);
                annuboamt1 = Convert.ToDouble(oForm.Items.Item("annuboamt1").Specific.VALUE);

                entno2 = oForm.Items.Item("entno2").Specific.VALUE.ToString().Trim();
                servcomp2 = oForm.Items.Item("servcomp2").Specific.VALUE.ToString().Trim();

                if (string.IsNullOrEmpty(oForm.Items.Item("symd2").Specific.VALUE.ToString().Trim()))
                {
                    symd2 = "";
                    eymd2 = "";
                }
                else
                {
                    symd2 = oForm.Items.Item("symd2").Specific.VALUE.ToString().Trim();
                    eymd2 = oForm.Items.Item("eymd2").Specific.VALUE.ToString().Trim();
                }
                
                payrtot2 = Convert.ToDouble(oForm.Items.Item("payrtot2").Specific.VALUE);
                bnstot2 = Convert.ToDouble(oForm.Items.Item("bnstot2").Specific.VALUE);
                fwork2 = Convert.ToDouble(oForm.Items.Item("fwork2").Specific.VALUE);
                ndtalw2 = Convert.ToDouble(oForm.Items.Item("ndtalw2").Specific.VALUE);
                etcntax2 = Convert.ToDouble(oForm.Items.Item("etcntax2").Specific.VALUE);
                lnchalw2 = Convert.ToDouble(oForm.Items.Item("lnchalw2").Specific.VALUE);
                ftaxamt2 = Convert.ToDouble(oForm.Items.Item("ftaxamt2").Specific.VALUE);
                savtaxddc2 = Convert.ToDouble(oForm.Items.Item("savtaxddc2").Specific.VALUE);
                indmtax2 = Convert.ToDouble(oForm.Items.Item("indmtax2").Specific.VALUE);
                fvsptax2 = Convert.ToDouble(oForm.Items.Item("fvsptax2").Specific.VALUE);
                residtax2 = Convert.ToDouble(oForm.Items.Item("residtax2").Specific.VALUE);
                medcinsr2 = Convert.ToDouble(oForm.Items.Item("medcinsr2").Specific.VALUE);
                asopinsr2 = Convert.ToDouble(oForm.Items.Item("asopinsr2").Specific.VALUE);
                annuboamt2 = Convert.ToDouble(oForm.Items.Item("annuboamt2").Specific.VALUE);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                YEAR = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.VALUE.ToString().Trim();


                if (string.IsNullOrWhiteSpace(CLTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(YEAR))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(MSTCOD))
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    //신규
                    sQry = "INSERT INTO [p_sbservcomp]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "entno1,";
                    sQry = sQry + "servcomp1,";
                    sQry = sQry + "symd1,";
                    sQry = sQry + "eymd1,";
                    sQry = sQry + "payrtot1,";
                    sQry = sQry + "bnstot1,";
                    sQry = sQry + "fwork1,";
                    sQry = sQry + "ndtalw1,";
                    sQry = sQry + "etcntax1,";
                    sQry = sQry + "lnchalw1,";
                    sQry = sQry + "ftaxamt1,";
                    sQry = sQry + "savtaxddc1,";
                    sQry = sQry + "incmtax1,";
                    sQry = sQry + "fvsptax1,";
                    sQry = sQry + "residtax1,";
                    sQry = sQry + "medcinsr1,";
                    sQry = sQry + "asopinsr1,";
                    sQry = sQry + "annuboamt1,";
                    sQry = sQry + "entno2,";
                    sQry = sQry + "servcomp2,";
                    sQry = sQry + "symd2,";
                    sQry = sQry + "eymd2,";
                    sQry = sQry + "payrtot2,";
                    sQry = sQry + "bnstot2,";
                    sQry = sQry + "fwork2,";
                    sQry = sQry + "ndtalw2,";
                    sQry = sQry + "etcntax2,";
                    sQry = sQry + "lnchalw2,";
                    sQry = sQry + "ftaxamt2,";
                    sQry = sQry + "savtaxddc2,";
                    sQry = sQry + "indmtax2,";
                    sQry = sQry + "fvsptax2,";
                    sQry = sQry + "residtax2,";
                    sQry = sQry + "medcinsr2,";
                    sQry = sQry + "asopinsr2,";
                    sQry = sQry + "annuboamt2,";
                    sQry = sQry + "jscntddc1,";
                    sQry = sQry + "jscntddc2";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + CLTCOD + "',";
                    sQry = sQry + "'" + YEAR + "',";
                    sQry = sQry + "'" + MSTCOD + "',";
                    sQry = sQry + "'" + entno1 + "',";
                    sQry = sQry + "'" + servcomp1 + "',";
                    sQry = sQry + "'" + symd1 + "',";
                    sQry = sQry + "'" + eymd1 + "',";

                    sQry = sQry + payrtot1 + ",";
                    sQry = sQry + bnstot1 + ",";
                    sQry = sQry + fwork1 + ",";
                    sQry = sQry + ndtalw1 + ",";
                    sQry = sQry + etcntax1 + ",";
                    sQry = sQry + lnchalw1 + ",";
                    sQry = sQry + ftaxamt1 + ",";
                    sQry = sQry + savtaxddc1 + ",";
                    sQry = sQry + incmtax1 + ",";
                    sQry = sQry + fvsptax1 + ",";
                    sQry = sQry + residtax1 + ",";
                    sQry = sQry + medcinsr1 + ",";
                    sQry = sQry + asopinsr1 + ",";
                    sQry = sQry + annuboamt1 + ",";
                    sQry = sQry + "'" + entno2 + "',";
                    sQry = sQry + "'" + servcomp2 + "',";
                    sQry = sQry + "'" + symd2 + "',";
                    sQry = sQry + "'" + eymd2 + "',";
                    sQry = sQry + payrtot2 + ",";
                    sQry = sQry + bnstot2 + ",";
                    sQry = sQry + fwork2 + ",";
                    sQry = sQry + ndtalw2 + ",";
                    sQry = sQry + etcntax2 + ",";
                    sQry = sQry + lnchalw2 + ",";
    ;               sQry = sQry + ftaxamt2 + ",";
                    sQry = sQry + savtaxddc2 + ",";
                    sQry = sQry + indmtax2 + ",";
                    sQry = sQry + fvsptax2 + ",";
                    sQry = sQry + residtax2 + ",";
                    sQry = sQry + medcinsr2 + ",";
                    sQry = sQry + asopinsr2 + ",";
                    sQry = sQry + annuboamt2 + ","; 
                    sQry = sQry + "0, 0 )";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY401_DataFind();
                }
                else
                {
                    //수정
                    sQry = "Update [p_sbservcomp] set ";
                    sQry = sQry + "entno1 = '" + entno1 + "',";
                    sQry = sQry + "servcomp1 = '" + servcomp1 + "',";
                    sQry = sQry + "symd1 = '" + symd1 + "',";
                    sQry = sQry + "eymd1 = '" + eymd1 + "',";
                    sQry = sQry + "payrtot1 = " + payrtot1 + ",";
                    sQry = sQry + "bnstot1 = " + bnstot1 + ",";
                    sQry = sQry + "fwork1 = " + fwork1 + ",";
                    sQry = sQry + "ndtalw1 = " + ndtalw1 + ",";
                    sQry = sQry + "etcntax1 = " + etcntax1 + ",";
                    sQry = sQry + "lnchalw1 = " + lnchalw1 + ",";
                    sQry = sQry + "ftaxamt1 = " + ftaxamt1 + ",";
                    sQry = sQry + "savtaxddc1 = " + savtaxddc1 + ",";
                    sQry = sQry + "incmtax1 = " + incmtax1 + ",";
                    sQry = sQry + "fvsptax1 = " + fvsptax1 + ",";
                    sQry = sQry + "residtax1 = " + residtax1 + ",";
                    sQry = sQry + "medcinsr1 = " + medcinsr1 + ",";
                    sQry = sQry + "asopinsr1 = " + asopinsr1 + ",";
                    sQry = sQry + "annuboamt1 =" + annuboamt1 + ",";
                    sQry = sQry + "entno2 = '" + entno2 + "',";
                    sQry = sQry + "servcomp2 = '" + servcomp2 + "',";
                    sQry = sQry + "symd2 = '" + symd2 + "',";
                    sQry = sQry + "eymd2 = '" + eymd2 + "',";
                    sQry = sQry + "payrtot2 = " + payrtot2 + ",";
                    sQry = sQry + "bnstot2= " + bnstot2 + ",";
                    sQry = sQry + "fwork2 = " + fwork2 + ",";
                    sQry = sQry + "ndtalw2 = " + ndtalw2 + ",";
                    sQry = sQry + "etcntax2 = " + etcntax2 + ",";
                    sQry = sQry + "lnchalw2 = " + lnchalw2 + ",";
                    sQry = sQry + "ftaxamt2 = " + ftaxamt2 + ",";
                    sQry = sQry + "savtaxddc2 = " + savtaxddc2 + ",";
                    sQry = sQry + "indmtax2 = " + indmtax2 + ",";
                    sQry = sQry + "fvsptax2 = " + fvsptax2 + ",";
                    sQry = sQry + "residtax2 = " + residtax2 + ",";
                    sQry = sQry + "medcinsr2 = " + medcinsr2 + ",";
                    sQry = sQry + "asopinsr2 = " + asopinsr2 + ",";
                    sQry = sQry + "annuboamt2 =" + annuboamt2;

                    sQry = sQry + " Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY401_DataFind();
                }
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장코드를 확인 하세요.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("년도를 확인 하세요.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("사원코드를 확인 하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }
    
        /// <summary>
        /// PH_PY401_Delete
        /// </summary>
        private void PH_PY401_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD, YEAR, MSTCOD = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                YEAR = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                
                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY401.Rows.Count > 0)
                    {
                        sQry = " Delete From [p_sbservcomp] Where  saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY401_DataFind();
                    }
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                //    PSH_Globals.SBO_Application.MessageBox("급여계산 된 자료는 삭제할 수 없습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY401_TitleSetting
        /// </summary>
        private void PH_PY401_TitleSetting(ref int iRow)
        {
            int i = 0;
            string[] COLNAM = new string[6];

            try
            {
                //그리드 콤보박스
                COLNAM[0] = "년도";
                COLNAM[1] = "부서";
                COLNAM[2] = "담당";
                COLNAM[3] = "사번";
                COLNAM[4] = "성명";
                COLNAM[5] = "직급";
                
                for (i = 0; i <= Information.UBound(COLNAM); i++)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY401_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
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
//	internal class PH_PY401
//	{
//////********************************************************************************
//////  File           : PH_PY401.cls
//////  Module         : 인사관리 > 연말정산관리
//////  Desc           : 전근무지등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Grid oGrid1;
//		public SAPbouiCOM.DataTable oDS_PH_PY401A;


//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY401.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY401_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY401");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "Code"

//			oForm.PaneLevel = 1;
//			oForm.Freeze(true);
//			PH_PY401_CreateItems();
//			PH_PY401_FormItemEnabled();
//			PH_PY401_EnableMenus();
//			//    Call PH_PY401_SetDocument(oFromDocEntry01)
//			//    Call PH_PY401_FormResize

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

//		private bool PH_PY401_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;
//			string CLTCOD = null;

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

//			oGrid1 = oForm.Items.Item("Grid01").Specific;

//			oForm.DataSources.DataTables.Add("PH_PY401");

//			oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY401");
//			oDS_PH_PY401A = oForm.DataSources.DataTables.Item("PH_PY401");


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			////사업장

//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			//
//			//    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + oCompany.UserName + "'")
//			//    oCombo.Select CLTCOD, psk_ByValue


//			//    oCombo.Select 0, psk_Index
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			////년도
//			oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

//			////사번
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
//			////성명
//			oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

//			////사업자등록번호1
//			oForm.DataSources.UserDataSources.Add("entno1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("entno1").Specific.DataBind.SetBound(true, "", "entno1");
//			////사업자등록번호2
//			oForm.DataSources.UserDataSources.Add("entno2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("entno2").Specific.DataBind.SetBound(true, "", "entno2");

//			////근무처명1
//			oForm.DataSources.UserDataSources.Add("servcomp1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("servcomp1").Specific.DataBind.SetBound(true, "", "servcomp1");
//			////근무처명2
//			oForm.DataSources.UserDataSources.Add("servcomp2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("servcomp2").Specific.DataBind.SetBound(true, "", "servcomp2");

//			////시작일1
//			oForm.DataSources.UserDataSources.Add("symd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("symd1").Specific.DataBind.SetBound(true, "", "symd1");
//			////시작일2
//			oForm.DataSources.UserDataSources.Add("symd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("symd2").Specific.DataBind.SetBound(true, "", "symd2");

//			////종료일1
//			oForm.DataSources.UserDataSources.Add("eymd1", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("eymd1").Specific.DataBind.SetBound(true, "", "eymd1");
//			////종료일2
//			oForm.DataSources.UserDataSources.Add("eymd2", SAPbouiCOM.BoDataType.dt_DATE, 8);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("eymd2").Specific.DataBind.SetBound(true, "", "eymd2");

//			////급여총액1
//			oForm.DataSources.UserDataSources.Add("payrtot1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("payrtot1").Specific.DataBind.SetBound(true, "", "payrtot1");
//			////급여총액2
//			oForm.DataSources.UserDataSources.Add("payrtot2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("payrtot2").Specific.DataBind.SetBound(true, "", "payrtot2");

//			////상여총액1
//			oForm.DataSources.UserDataSources.Add("bnstot1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("bnstot1").Specific.DataBind.SetBound(true, "", "bnstot1");
//			////상여총액2
//			oForm.DataSources.UserDataSources.Add("bnstot2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("bnstot2").Specific.DataBind.SetBound(true, "", "bnstot2");

//			////국외근로1
//			oForm.DataSources.UserDataSources.Add("fwork1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("fwork1").Specific.DataBind.SetBound(true, "", "fwork1");
//			////국외근로2
//			oForm.DataSources.UserDataSources.Add("fwork2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("fwork2").Specific.DataBind.SetBound(true, "", "fwork2");

//			////야간근로수당등1
//			oForm.DataSources.UserDataSources.Add("ndtalw1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ndtalw1").Specific.DataBind.SetBound(true, "", "ndtalw1");
//			////야간근로수당등2
//			oForm.DataSources.UserDataSources.Add("ndtalw2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ndtalw2").Specific.DataBind.SetBound(true, "", "ndtalw2");

//			////기타비과세1
//			oForm.DataSources.UserDataSources.Add("etcntax1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("etcntax1").Specific.DataBind.SetBound(true, "", "etcntax1");
//			////기타비과세2
//			oForm.DataSources.UserDataSources.Add("etcntax2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("etcntax2").Specific.DataBind.SetBound(true, "", "etcntax2");

//			////중식수당1
//			oForm.DataSources.UserDataSources.Add("lnchalw1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("lnchalw1").Specific.DataBind.SetBound(true, "", "lnchalw1");
//			////중식수당2
//			oForm.DataSources.UserDataSources.Add("lnchalw2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("lnchalw2").Specific.DataBind.SetBound(true, "", "lnchalw2");

//			////외국납부세액1
//			oForm.DataSources.UserDataSources.Add("ftaxamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ftaxamt1").Specific.DataBind.SetBound(true, "", "ftaxamt1");
//			////외국납부세액2
//			oForm.DataSources.UserDataSources.Add("ftaxamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ftaxamt2").Specific.DataBind.SetBound(true, "", "ftaxamt2");

//			////저축세액공제1
//			oForm.DataSources.UserDataSources.Add("savtaxddc1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("savtaxddc1").Specific.DataBind.SetBound(true, "", "savtaxddc1");
//			////저축세액공제2
//			oForm.DataSources.UserDataSources.Add("savtaxddc2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("savtaxddc2").Specific.DataBind.SetBound(true, "", "savtaxddc2");

//			////소득세1
//			oForm.DataSources.UserDataSources.Add("incmtax1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("incmtax1").Specific.DataBind.SetBound(true, "", "incmtax1");
//			////소득세2
//			oForm.DataSources.UserDataSources.Add("indmtax2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("indmtax2").Specific.DataBind.SetBound(true, "", "indmtax2");

//			////농어촌특별세1
//			oForm.DataSources.UserDataSources.Add("fvsptax1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("fvsptax1").Specific.DataBind.SetBound(true, "", "fvsptax1");
//			////농어촌특별세2
//			oForm.DataSources.UserDataSources.Add("fvsptax2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("fvsptax2").Specific.DataBind.SetBound(true, "", "fvsptax2");

//			////주민세1
//			oForm.DataSources.UserDataSources.Add("residtax1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("residtax1").Specific.DataBind.SetBound(true, "", "residtax1");
//			////주민세2
//			oForm.DataSources.UserDataSources.Add("residtax2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("residtax2").Specific.DataBind.SetBound(true, "", "residtax2");

//			////건강보험1
//			oForm.DataSources.UserDataSources.Add("medcinsr1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("medcinsr1").Specific.DataBind.SetBound(true, "", "medcinsr1");
//			////건강보험2
//			oForm.DataSources.UserDataSources.Add("medcinsr2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("medcinsr2").Specific.DataBind.SetBound(true, "", "medcinsr2");

//			////고용보험1
//			oForm.DataSources.UserDataSources.Add("asopinsr1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("asopinsr1").Specific.DataBind.SetBound(true, "", "asopinsr1");
//			////고용보험2
//			oForm.DataSources.UserDataSources.Add("asopinsr2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("asopinsr2").Specific.DataBind.SetBound(true, "", "asopinsr2");

//			////연금보험료공제1
//			oForm.DataSources.UserDataSources.Add("annuboamt1", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("annuboamt1").Specific.DataBind.SetBound(true, "", "annuboamt1");
//			////고용보험2
//			oForm.DataSources.UserDataSources.Add("annuboamt2", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("annuboamt2").Specific.DataBind.SetBound(true, "", "annuboamt2");


//			oForm.Update();

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
//			PH_PY401_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY401_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", false);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", false);
//			////행삭제

//			return;
//			PH_PY401_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY401_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY401_FormItemEnabled();
//				//        Call PH_PY401_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY401_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY401_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY401_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			int i = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;


//			string CLTCOD = null;
//			string sPosDate = null;

//			 // ERROR: Not supported in C#: OnErrorStatement

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {


//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//				//UPGRADE_WARNING: oForm.Items(Year).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Year").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.VALUE = "";
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("FullName").Specific.VALUE = "";

//				oForm.DataSources.UserDataSources.Item("entno1").Value = "";
//				oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
//				oForm.DataSources.UserDataSources.Item("symd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

//				oForm.DataSources.UserDataSources.Item("payrtot1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("bnstot1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fwork1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ndtalw1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("etcntax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("lnchalw1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("incmtax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fvsptax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("residtax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("medcinsr1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("asopinsr1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("annuboamt1").Value = Convert.ToString(0);


//				oForm.DataSources.UserDataSources.Item("entno2").Value = "";
//				oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
//				oForm.DataSources.UserDataSources.Item("symd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

//				oForm.DataSources.UserDataSources.Item("payrtot2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("bnstot2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fwork2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ndtalw2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("etcntax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("lnchalw2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("indmtax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fvsptax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("residtax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("medcinsr2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("asopinsr2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("annuboamt2").Value = Convert.ToString(0);

//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("TeamName").Specific.VALUE = "";
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("RspName").Specific.VALUE = "";
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("ClsName").Specific.VALUE = "";

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");


//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가


//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY401_FormItemEnabled_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			string tSex = null;
//			string tBrith = null;
//			//UPGRADE_NOTE: Day이(가) Day_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string Day_Renamed = null;
//			string ActCode = null;
//			string CLTCOD = null;
//			string MSTCOD = null;
//			string yyyy = null;
//			string Result = null;

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (PH_PY401_DataValidCheck() == false) {
//								BubbleEvent = false;
//							}
//						}

//						if (pval.ItemUID == "Btn_ret") {
//							PH_PY401_MTX01();
//						}



//						if (pval.ItemUID == "Btn01") {

//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							yyyy = oForm.Items.Item("Year").Specific.VALUE;
//							sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
//							oRecordSet.DoQuery(sQry);

//							//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							Result = oRecordSet.Fields.Item(0).Value;
//							if (Result != "Y") {
//								MDC_Globals.Sbo_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
//							}
//							if (Result == "Y") {
//								PH_PY401_SAVE();
//							}
//						}


//						if (pval.ItemUID == "Btn_del") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							yyyy = oForm.Items.Item("Year").Specific.VALUE;
//							sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
//							oRecordSet.DoQuery(sQry);

//							//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							Result = oRecordSet.Fields.Item(0).Value;
//							if (Result != "Y") {
//								MDC_Globals.Sbo_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
//							}
//							if (Result == "Y") {
//								PH_PY401_Delete();
//								PH_PY401_FormItemEnabled();
//							}
//						}
//						//                If oForm.Mode = fm_FIND_MODE Then
//						//                    If pval.ItemUID = "Btn01" Then
//						//                        Sbo_Application.ActivateMenuItem ("7425")
//						//                        BubbleEvent = False
//						//                    End If
//						//
//						//                End If
//					} else if (pval.BeforeAction == false) {
//						switch (pval.ItemUID) {
//							case "1":
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY401_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY401_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY401_FormItemEnabled();
//									}
//								}
//								break;
//							//
//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true) {
//						if (pval.CharPressed == 9) {
//							if (pval.ItemUID == "MSTCOD") {
//								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							}
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
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
//							////사업장(헤더)
//							if (pval.ItemUID == "SCLTCOD") {

//							}

//						}
//					}

//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row >= 0) {
//									switch (pval.ItemUID) {
//										case "Grid01":
//											//Call oMat1.SelectRow(pval.Row, True, False)
//											PH_PY401_MTX02(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//											break;
//									}

//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Grid01":
//								if (pval.Row > 0) {
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
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {
//					} else {

//					}
//					oForm.Freeze(false);
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
//					//            Call oForm.Freeze(True)
//					if (pval.BeforeAction == true) {
//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							switch (pval.ItemUID) {

//								case "MSTCOD":
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

//									sQry = "Select Code,";
//									sQry = sQry + " FullName = U_FullName,";
//									sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '1'";
//									sQry = sQry + " And U_Code = U_TeamCode),''),";
//									sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '2'";
//									sQry = sQry + " And U_Code = U_RspCode),''),";
//									sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm";
//									sQry = sQry + " From [@PS_HR200L]";
//									sQry = sQry + " WHERE Code = '9'";
//									sQry = sQry + " And U_Code  = U_ClsCode";
//									sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
//									sQry = sQry + " From [@PH_PY001A]";
//									sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
//									sQry = sQry + " and Code = '" + MSTCOD + "'";

//									oRecordSet.DoQuery(sQry);

//									//UPGRADE_WARNING: oForm.Items(FullName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
//									//UPGRADE_WARNING: oForm.Items(TeamName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
//									//UPGRADE_WARNING: oForm.Items(RspName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
//									//UPGRADE_WARNING: oForm.Items(ClsName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
//									break;

//							}

//						}
//					}
//					break;
//				//            Call oForm.Freeze(False)
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						//                oMat1.LoadFromDataSource
//						//                Call PH_PY401_AddMatrixRow

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
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY401A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY401A = null;

//						//                Set oMat1 = Nothing
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
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//					}
//					break;
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
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY401A", "Code")
//						//                End If
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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//					//                Call PH_PY401_FormItemEnabled
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY401_FormItemEnabled();
//						break;
//					//                Call PH_PY401_AddMatrixRow
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY401_FormItemEnabled();
//						//                Call PH_PY401_AddMatrixRow
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY401_FormItemEnabled();
//						break;
//					//                Call PH_PY401_AddMatrixRow
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY401_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			int i = 0;
//			string sQry = null;
//			SAPbouiCOM.ComboBox oCombo = null;

//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Raise_FormDataEvent_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Mat01":
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


//		public void PH_PY401_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY401'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY401_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY401_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			int j = 0;

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY401_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY401_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;
//			int iRow = 0;

//			string Param01 = null;
//			string Param02 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Year").Specific.VALUE;

//			if (string.IsNullOrEmpty(Strings.Trim(Param01))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY401_MTX01_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Param02))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY401_MTX01_Exit;
//			}



//			sQry = "EXEC PH_PY401_01 '" + Param01 + "', '" + Param02 + "'";

//			oDS_PH_PY401A.ExecuteQuery(sQry);



//			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

//			PH_PY401_TitleSetting(ref iRow);

//			oForm.Update();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY401_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY401_MTX01_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY401_MTX02(string oUID, ref int oRow = 0, ref string oCol = "")
//		{


//			////그리드 자료를 head에 로드

//			int i = 0;
//			string sQry = null;
//			int sRow = 0;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sRow = oRow;


//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oDS_PH_PY401A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oDS_PH_PY401A.Columns.Item("Year").Cells.Item(oRow).Value;
//			//UPGRADE_WARNING: oDS_PH_PY401A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oDS_PH_PY401A.Columns.Item("MSTCOD").Cells.Item(oRow).Value;



//			sQry = "EXEC PH_PY401_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
//			oRecordSet.DoQuery(sQry);

//			if ((oRecordSet.RecordCount == 0)) {

//				//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oDS_PH_PY401A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("MSTCOD").Specific.VALUE = oDS_PH_PY401A.Columns.Item("MSTCOD").Cells.Item(oRow).Value;
//				//UPGRADE_WARNING: oForm.Items(FullName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: oDS_PH_PY401A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("FullName").Specific.VALUE = oDS_PH_PY401A.Columns.Item("FullName").Cells.Item(oRow).Value;

//				oForm.DataSources.UserDataSources.Item("entno1").Value = "";
//				oForm.DataSources.UserDataSources.Item("servcomp1").Value = "";
//				oForm.DataSources.UserDataSources.Item("symd1").Value = "";
//				oForm.DataSources.UserDataSources.Item("eymd1").Value = "";

//				oForm.DataSources.UserDataSources.Item("payrtot1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("bnstot1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fwork1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ndtalw1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("etcntax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("lnchalw1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("incmtax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fvsptax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("residtax1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("medcinsr1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("asopinsr1").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("annuboamt1").Value = Convert.ToString(0);


//				oForm.DataSources.UserDataSources.Item("entno2").Value = "";
//				oForm.DataSources.UserDataSources.Item("servcomp2").Value = "";
//				oForm.DataSources.UserDataSources.Item("symd2").Value = "";
//				oForm.DataSources.UserDataSources.Item("eymd2").Value = "";

//				oForm.DataSources.UserDataSources.Item("payrtot2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("bnstot2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fwork2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ndtalw2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("etcntax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("lnchalw2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("indmtax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("fvsptax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("residtax2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("medcinsr2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("asopinsr2").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("annuboamt2").Value = Convert.ToString(0);



//				//        oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
//				//        oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
//				//        oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE

//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY401_MTX02_Exit;
//			}


//			//    oForm.Items("PosDate").Specific.Value = Format(oRecordset.Fields("PosDate").Value, "YYYYMMDD")

//			//    oForm.DataSources.UserDataSources.Item("PosDate").Value = Format(oRecordSet.Fields("PosDate").Value, "YYYYMMDD")


//			//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTCOD").Specific.VALUE = oRecordSet.Fields.Item("MSTCOD").Value;
//			//UPGRADE_WARNING: oForm.Items(FullName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("entno1").Value = oRecordSet.Fields.Item("entno1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("servcomp1").Value = oRecordSet.Fields.Item("servcomp1").Value;
//			oForm.DataSources.UserDataSources.Item("symd1").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("symd1").Value, "YYYY-MM-DD");
//			oForm.DataSources.UserDataSources.Item("eymd1").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("eymd1").Value, "YYYY-MM-DD");

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("payrtot1").Value = oRecordSet.Fields.Item("payrtot1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("bnstot1").Value = oRecordSet.Fields.Item("bnstot1").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("fwork1").Value = oRecordSet.Fields.Item("fwork1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ndtalw1").Value = oRecordSet.Fields.Item("ndtalw1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("etcntax1").Value = oRecordSet.Fields.Item("etcntax1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("lnchalw1").Value = oRecordSet.Fields.Item("lnchalw1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ftaxamt1").Value = oRecordSet.Fields.Item("ftaxamt1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("savtaxddc1").Value = oRecordSet.Fields.Item("savtaxddc1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("incmtax1").Value = oRecordSet.Fields.Item("incmtax1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("residtax1").Value = oRecordSet.Fields.Item("residtax1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("medcinsr1").Value = oRecordSet.Fields.Item("medcinsr1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("asopinsr1").Value = oRecordSet.Fields.Item("asopinsr1").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("annuboamt1").Value = oRecordSet.Fields.Item("annuboamt1").Value;


//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("entno2").Value = oRecordSet.Fields.Item("entno2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("servcomp2").Value = oRecordSet.Fields.Item("servcomp2").Value;
//			oForm.DataSources.UserDataSources.Item("symd2").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("symd2").Value, "YYYY-MM-DD");
//			oForm.DataSources.UserDataSources.Item("eymd2").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("eymd2").Value, "YYYY-MM-DD");

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("payrtot2").Value = oRecordSet.Fields.Item("payrtot2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("bnstot2").Value = oRecordSet.Fields.Item("bnstot2").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("fwork2").Value = oRecordSet.Fields.Item("fwork2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ndtalw2").Value = oRecordSet.Fields.Item("ndtalw2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("etcntax2").Value = oRecordSet.Fields.Item("etcntax2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("lnchalw2").Value = oRecordSet.Fields.Item("lnchalw2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ftaxamt2").Value = oRecordSet.Fields.Item("ftaxamt2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("savtaxddc2").Value = oRecordSet.Fields.Item("savtaxddc2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("indmtax2").Value = oRecordSet.Fields.Item("indmtax2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("residtax2").Value = oRecordSet.Fields.Item("residtax2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("medcinsr2").Value = oRecordSet.Fields.Item("medcinsr2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("asopinsr2").Value = oRecordSet.Fields.Item("asopinsr2").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("annuboamt2").Value = oRecordSet.Fields.Item("annuboamt2").Value;


//			//    '//부서
//			//UPGRADE_WARNING: oForm.Items(TeamName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
//			//UPGRADE_WARNING: oForm.Items(RspName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
//			//UPGRADE_WARNING: oForm.Items(ClsName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;



//			oForm.Update();


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY401_MTX02_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY401_MTX02_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY401_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY401A] WHERE DocEntry = ' + oForm.Items(DocEntry).Specific.VALUE + ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY401A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY401_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY401_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY401_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//////행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent, ref SAPbouiCOM.Matrix oMat, ref SAPbouiCOM.DBDataSource DBData, ref string CheckField)
//		{

//			int i = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((oLastColRow > 0)) {
//				if (pval.BeforeAction == true) {

//				} else if (pval.BeforeAction == false) {
//					if (oMat.RowCount != oMat.VisualRowCount) {
//						oMat.FlushToDataSource();

//						while ((i <= DBData.Size - 1)) {
//							if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i))) {
//								DBData.RemoveRecord((i));
//								i = 0;
//							} else {
//								i = i + 1;
//							}
//						}

//						for (i = 0; i <= DBData.Size; i++) {
//							DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//						}

//						oMat.LoadFromDataSource();
//					}
//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY401_SAVE()
//		{

//			////데이타 저장

//			int i = 0;
//			string sQry = null;

//			//UPGRADE_NOTE: YEAR이(가) YEAR_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string FullName = null;
//			string CLTCOD = null;
//			string MSTCOD = null;
//			string YEAR_Renamed = null;



//			string symd1 = null;
//			string entno1 = null;
//			string servcomp1 = null;
//			string eymd1 = null;
//			string symd2 = null;
//			string entno2 = null;
//			string servcomp2 = null;
//			string eymd2 = null;

//			object medcinsr1 = null;
//			object fvsptax1 = null;
//			object savtaxddc1 = null;
//			object lnchalw1 = null;
//			object ndtalw1 = null;
//			object bnstot1 = null;
//			object payrtot1 = null;
//			object fwork1 = null;
//			object etcntax1 = null;
//			object ftaxamt1 = null;
//			object incmtax1 = null;
//			object residtax1 = null;
//			object asopinsr1 = null;
//			double annuboamt1 = 0;
//			object medcinsr2 = null;
//			object fvsptax2 = null;
//			object savtaxddc2 = null;
//			object lnchalw2 = null;
//			object ndtalw2 = null;
//			object bnstot2 = null;
//			object payrtot2 = null;
//			object fwork2 = null;
//			object etcntax2 = null;
//			object ftaxamt2 = null;
//			object indmtax2 = null;
//			object residtax2 = null;
//			object asopinsr2 = null;
//			double annuboamt2 = 0;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			entno1 = oForm.Items.Item("entno1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			servcomp1 = oForm.Items.Item("servcomp1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			symd1 = Strings.Trim(oForm.Items.Item("symd1").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			eymd1 = Strings.Trim(oForm.Items.Item("eymd1").Specific.VALUE);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: payrtot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			payrtot1 = oForm.Items.Item("payrtot1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: bnstot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			bnstot1 = oForm.Items.Item("bnstot1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: fwork1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			fwork1 = oForm.Items.Item("fwork1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ndtalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ndtalw1 = oForm.Items.Item("ndtalw1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: etcntax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			etcntax1 = oForm.Items.Item("etcntax1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: lnchalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			lnchalw1 = oForm.Items.Item("lnchalw1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ftaxamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ftaxamt1 = oForm.Items.Item("ftaxamt1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: savtaxddc1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			savtaxddc1 = oForm.Items.Item("savtaxddc1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: incmtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			incmtax1 = oForm.Items.Item("incmtax1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: fvsptax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			fvsptax1 = oForm.Items.Item("fvsptax1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: residtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			residtax1 = oForm.Items.Item("residtax1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: medcinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			medcinsr1 = oForm.Items.Item("medcinsr1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: asopinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			asopinsr1 = oForm.Items.Item("asopinsr1").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			annuboamt1 = oForm.Items.Item("annuboamt1").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			entno2 = oForm.Items.Item("entno2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			servcomp2 = oForm.Items.Item("servcomp2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items(symd2).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(oForm.Items.Item("symd2").Specific.VALUE)) {
//				symd2 = "";
//				eymd2 = "";
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				symd2 = Strings.Trim(oForm.Items.Item("symd2").Specific.VALUE);
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				eymd2 = Strings.Trim(oForm.Items.Item("eymd2").Specific.VALUE);
//			}


//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: payrtot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			payrtot2 = oForm.Items.Item("payrtot2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: bnstot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			bnstot2 = oForm.Items.Item("bnstot2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: fwork2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			fwork2 = oForm.Items.Item("fwork2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ndtalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ndtalw2 = oForm.Items.Item("ndtalw2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: etcntax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			etcntax2 = oForm.Items.Item("etcntax2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: lnchalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			lnchalw2 = oForm.Items.Item("lnchalw2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ftaxamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ftaxamt2 = oForm.Items.Item("ftaxamt2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: savtaxddc2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			savtaxddc2 = oForm.Items.Item("savtaxddc2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: indmtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			indmtax2 = oForm.Items.Item("indmtax2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: fvsptax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			fvsptax2 = oForm.Items.Item("fvsptax2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: residtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			residtax2 = oForm.Items.Item("residtax2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: medcinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			medcinsr2 = oForm.Items.Item("medcinsr2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: asopinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			asopinsr2 = oForm.Items.Item("asopinsr2").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			annuboamt2 = oForm.Items.Item("annuboamt2").Specific.VALUE;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FullName = oForm.Items.Item("FullName").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			YEAR_Renamed = oForm.Items.Item("Year").Specific.VALUE;

//			if (string.IsNullOrEmpty(Strings.Trim(YEAR_Renamed))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY401_SAVE_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(CLTCOD))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY401_SAVE_Exit;
//			}
//			if (string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY401_SAVE_Exit;
//			}

//			sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR_Renamed + "' And sabun = '" + MSTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.Fields.Item(0).Value > 0) {
//				////갱신

//				sQry = "Update [p_sbservcomp] set ";
//				sQry = sQry + "entno1 = '" + entno1 + "',";
//				sQry = sQry + "servcomp1 = '" + servcomp1 + "',";
//				sQry = sQry + "symd1 = '" + symd1 + "',";
//				sQry = sQry + "eymd1 = '" + eymd1 + "',";
//				//UPGRADE_WARNING: payrtot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "payrtot1 = " + payrtot1 + ",";
//				//UPGRADE_WARNING: bnstot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "bnstot1 = " + bnstot1 + ",";
//				//UPGRADE_WARNING: fwork1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "fwork1 = " + fwork1 + ",";
//				//UPGRADE_WARNING: ndtalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ndtalw1 = " + ndtalw1 + ",";
//				//UPGRADE_WARNING: etcntax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "etcntax1 = " + etcntax1 + ",";
//				//UPGRADE_WARNING: lnchalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "lnchalw1 = " + lnchalw1 + ",";
//				//UPGRADE_WARNING: ftaxamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ftaxamt1 = " + ftaxamt1 + ",";
//				//UPGRADE_WARNING: savtaxddc1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "savtaxddc1 = " + savtaxddc1 + ",";
//				//UPGRADE_WARNING: incmtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "incmtax1 = " + incmtax1 + ",";
//				//UPGRADE_WARNING: fvsptax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "fvsptax1 = " + fvsptax1 + ",";
//				//UPGRADE_WARNING: residtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "residtax1 = " + residtax1 + ",";
//				//UPGRADE_WARNING: medcinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "medcinsr1 = " + medcinsr1 + ",";
//				//UPGRADE_WARNING: asopinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "asopinsr1 = " + asopinsr1 + ",";
//				sQry = sQry + "annuboamt1 =" + annuboamt1 + ",";
//				sQry = sQry + "entno2 = '" + entno2 + "',";
//				sQry = sQry + "servcomp2 = '" + servcomp2 + "',";
//				sQry = sQry + "symd2 = '" + symd2 + "',";
//				sQry = sQry + "eymd2 = '" + eymd2 + "',";
//				//UPGRADE_WARNING: payrtot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "payrtot2 = " + payrtot2 + ",";
//				//UPGRADE_WARNING: bnstot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "bnstot2= " + bnstot2 + ",";
//				//UPGRADE_WARNING: fwork2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "fwork2 = " + fwork2 + ",";
//				//UPGRADE_WARNING: ndtalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ndtalw2 = " + ndtalw2 + ",";
//				//UPGRADE_WARNING: etcntax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "etcntax2 = " + etcntax2 + ",";
//				//UPGRADE_WARNING: lnchalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "lnchalw2 = " + lnchalw2 + ",";
//				//UPGRADE_WARNING: ftaxamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ftaxamt2 = " + ftaxamt2 + ",";
//				//UPGRADE_WARNING: savtaxddc2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "savtaxddc2 = " + savtaxddc2 + ",";
//				//UPGRADE_WARNING: indmtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "indmtax2 = " + indmtax2 + ",";
//				//UPGRADE_WARNING: fvsptax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "fvsptax2 = " + fvsptax2 + ",";
//				//UPGRADE_WARNING: residtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "residtax2 = " + residtax2 + ",";
//				//UPGRADE_WARNING: medcinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "medcinsr2 = " + medcinsr2 + ",";
//				//UPGRADE_WARNING: asopinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "asopinsr2 = " + asopinsr2 + ",";
//				sQry = sQry + "annuboamt2 =" + annuboamt2;

//				sQry = sQry + " Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR_Renamed + "' And sabun = '" + MSTCOD + "'";

//				oRecordSet.DoQuery(sQry);

//			} else {

//				////신규
//				sQry = "INSERT INTO [p_sbservcomp]";
//				sQry = sQry + " (";
//				sQry = sQry + "saup,";
//				sQry = sQry + "yyyy,";
//				sQry = sQry + "sabun,";
//				sQry = sQry + "entno1,";
//				sQry = sQry + "servcomp1,";
//				sQry = sQry + "symd1,";
//				sQry = sQry + "eymd1,";
//				sQry = sQry + "payrtot1,";
//				sQry = sQry + "bnstot1,";
//				sQry = sQry + "fwork1,";
//				sQry = sQry + "ndtalw1,";
//				sQry = sQry + "etcntax1,";
//				sQry = sQry + "lnchalw1,";
//				sQry = sQry + "ftaxamt1,";
//				sQry = sQry + "savtaxddc1,";
//				sQry = sQry + "incmtax1,";
//				sQry = sQry + "fvsptax1,";
//				sQry = sQry + "residtax1,";
//				sQry = sQry + "medcinsr1,";
//				sQry = sQry + "asopinsr1,";
//				sQry = sQry + "annuboamt1,";
//				sQry = sQry + "entno2,";
//				sQry = sQry + "servcomp2,";
//				sQry = sQry + "symd2,";
//				sQry = sQry + "eymd2,";
//				sQry = sQry + "payrtot2,";
//				sQry = sQry + "bnstot2,";
//				sQry = sQry + "fwork2,";
//				sQry = sQry + "ndtalw2,";
//				sQry = sQry + "etcntax2,";
//				sQry = sQry + "lnchalw2,";
//				sQry = sQry + "ftaxamt2,";
//				sQry = sQry + "savtaxddc2,";
//				sQry = sQry + "indmtax2,";
//				sQry = sQry + "fvsptax2,";
//				sQry = sQry + "residtax2,";
//				sQry = sQry + "medcinsr2,";
//				sQry = sQry + "asopinsr2,";
//				sQry = sQry + "annuboamt2,";
//				sQry = sQry + "jscntddc1,";
//				sQry = sQry + "jscntddc2";
//				sQry = sQry + " ) ";
//				sQry = sQry + "VALUES(";

//				sQry = sQry + "'" + CLTCOD + "',";
//				sQry = sQry + "'" + YEAR_Renamed + "',";
//				sQry = sQry + "'" + MSTCOD + "',";
//				sQry = sQry + "'" + entno1 + "',";
//				sQry = sQry + "'" + servcomp1 + "',";
//				sQry = sQry + "'" + symd1 + "',";
//				sQry = sQry + "'" + eymd1 + "',";

//				//UPGRADE_WARNING: payrtot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + payrtot1 + ",";
//				//UPGRADE_WARNING: bnstot1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + bnstot1 + ",";
//				//UPGRADE_WARNING: fwork1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + fwork1 + ",";
//				//UPGRADE_WARNING: ndtalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ndtalw1 + ",";
//				//UPGRADE_WARNING: etcntax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + etcntax1 + ",";
//				//UPGRADE_WARNING: lnchalw1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + lnchalw1 + ",";
//				//UPGRADE_WARNING: ftaxamt1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ftaxamt1 + ",";
//				//UPGRADE_WARNING: savtaxddc1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + savtaxddc1 + ",";
//				//UPGRADE_WARNING: incmtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + incmtax1 + ",";
//				//UPGRADE_WARNING: fvsptax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + fvsptax1 + ",";
//				//UPGRADE_WARNING: residtax1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + residtax1 + ",";
//				//UPGRADE_WARNING: medcinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + medcinsr1 + ",";
//				//UPGRADE_WARNING: asopinsr1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + asopinsr1 + ",";
//				sQry = sQry + annuboamt1 + ",";
//				sQry = sQry + "'" + entno2 + "',";
//				sQry = sQry + "'" + servcomp2 + "',";
//				sQry = sQry + "'" + symd2 + "',";
//				sQry = sQry + "'" + eymd2 + "',";
//				//UPGRADE_WARNING: payrtot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + payrtot2 + ",";
//				//UPGRADE_WARNING: bnstot2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + bnstot2 + ",";
//				//UPGRADE_WARNING: fwork2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + fwork2 + ",";
//				//UPGRADE_WARNING: ndtalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ndtalw2 + ",";
//				//UPGRADE_WARNING: etcntax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + etcntax2 + ",";
//				//UPGRADE_WARNING: lnchalw2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + lnchalw2 + ",";
//				//UPGRADE_WARNING: ftaxamt2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ftaxamt2 + ",";
//				//UPGRADE_WARNING: savtaxddc2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + savtaxddc2 + ",";
//				//UPGRADE_WARNING: indmtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + indmtax2 + ",";
//				//UPGRADE_WARNING: fvsptax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + fvsptax2 + ",";
//				//UPGRADE_WARNING: residtax2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + residtax2 + ",";
//				//UPGRADE_WARNING: medcinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + medcinsr2 + ",";
//				//UPGRADE_WARNING: asopinsr2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + asopinsr2 + ",";
//				sQry = sQry + annuboamt2 + ",";
//				sQry = sQry + "0, 0 )";

//				oRecordSet.DoQuery(sQry);
//			}


//			PH_PY401_FormItemEnabled();


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			PH_PY401_MTX01();

//			return;
//			PH_PY401_SAVE_Exit:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY401_SAVE_Error:
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY401_Delete()
//		{
//			////선택된 자료 삭제

//			string CLTCOD = null;
//			string MSTCOD = null;
//			//UPGRADE_NOTE: YEAR이(가) YEAR_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//			string YEAR_Renamed = null;
//			string FullName = null;


//			short i = 0;
//			short cnt = 0;

//			string sQry = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



//			oForm.Freeze(true);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			YEAR_Renamed = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FullName = oForm.Items.Item("FullName").Specific.VALUE;

//			sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR_Renamed + "' And sabun = '" + MSTCOD + "'";
//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			cnt = oRecordSet.Fields.Item(0).Value;
//			if (cnt > 0) {

//				if (string.IsNullOrEmpty(Strings.Trim(YEAR_Renamed))) {
//					MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY401_Delete_Exit;
//				}

//				if (string.IsNullOrEmpty(Strings.Trim(CLTCOD))) {
//					MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY401_Delete_Exit;
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(MSTCOD))) {
//					MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY401_Delete_Exit;
//				}




//				if (MDC_Globals.Sbo_Application.MessageBox(" 선택한사원('" + FullName + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1")) {
//					sQry = "Delete From [p_sbservcomp] Where saup = '" + CLTCOD + "' AND  yyyy = '" + YEAR_Renamed + "' And sabun = '" + MSTCOD + "' ";
//					oRecordSet.DoQuery(sQry);
//				}
//			}


//			oForm.Freeze(false);


//			PH_PY401_MTX01();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;


//			return;
//			PH_PY401_Delete_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			return;
//			PH_PY401_Delete_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY401_TitleSetting(ref int iRow)
//		{
//			int i = 0;
//			int j = 0;
//			string sQry = null;

//			string[] COLNAM = new string[6];

//			SAPbouiCOM.EditTextColumn oColumn = null;
//			SAPbouiCOM.ComboBoxColumn oComboCol = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			COLNAM[0] = "년도";
//			COLNAM[1] = "부서";
//			COLNAM[2] = "담당";
//			COLNAM[3] = "사번";
//			COLNAM[4] = "성명";
//			COLNAM[5] = "직급";

//			for (i = 0; i <= Information.UBound(COLNAM); i++) {
//				oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
//				oGrid1.Columns.Item(i).Editable = false;

//				oGrid1.Columns.Item(i).RightJustified = true;

//			}

//			oGrid1.AutoResizeColumns();

//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;

//			return;
//			Error_Message:

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY401_TitleSetting Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//	}
//}
