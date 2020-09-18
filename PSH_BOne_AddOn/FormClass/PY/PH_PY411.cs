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
    /// 연금저축등소득공제등록
    /// </summary>
    internal class PH_PY411 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY411;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY411.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY411_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY411");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY411_CreateItems();
                PH_PY411_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY411_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY411");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY411");
                oDS_PH_PY411 = oForm.DataSources.DataTables.Item("PH_PY411");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("순번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("구분코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("구분명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("금융기관코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("금융기관명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("계좌번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("납입년차", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("납입금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("공제금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("투자년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY411").Columns.Add("투자구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

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

                // 순번
                oForm.DataSources.UserDataSources.Add("seqn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("seqn").Specific.DataBind.SetBound(true, "", "seqn");

                // 공제구분
                oForm.DataSources.UserDataSources.Add("gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("gubun").Specific.DataBind.SetBound(true, "", "gubun");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '77' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("gubun").Specific, "Y");
                oForm.Items.Item("gubun").DisplayDesc = true;
                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 투자년도(중소기업창업투자조합출자) 18년추가
                oForm.DataSources.UserDataSources.Add("tyyyy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("tyyyy").Specific.DataBind.SetBound(true, "", "tyyyy");

                // 투자구분(중소기업창업투자조합출자) 18년추가
                oForm.DataSources.UserDataSources.Add("tgubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("tgubun").Specific.DataBind.SetBound(true, "", "tgubun");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '83' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("tgubun").Specific, "Y");
                oForm.Items.Item("tgubun").DisplayDesc = true;
                oForm.Items.Item("tgubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 금융기관코드
                oForm.DataSources.UserDataSources.Add("bcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("bcode").Specific.DataBind.SetBound(true, "", "bcode");

                // 금융기관명
                oForm.DataSources.UserDataSources.Add("bname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("bname").Specific.DataBind.SetBound(true, "", "bname");

                // 계좌번호
                oForm.DataSources.UserDataSources.Add("bnum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("bnum").Specific.DataBind.SetBound(true, "", "bnum");

                // 납입년차
                oForm.DataSources.UserDataSources.Add("yuncha", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("yuncha").Specific.DataBind.SetBound(true, "", "yuncha");

                // 불입금액
                oForm.DataSources.UserDataSources.Add("amt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("amt").Specific.DataBind.SetBound(true, "", "amt");

                // 공제금액
                oForm.DataSources.UserDataSources.Add("gamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("gamt").Specific.DataBind.SetBound(true, "", "gamt");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY411_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oForm.EnableMenu("1282", true);      // 문서추가

                if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE)))
                {
                    oForm.Items.Item("Year").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);
                }
                oForm.Items.Item("seqn").Specific.VALUE = "";
                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("tyyyy").Specific.VALUE = "";
                oForm.Items.Item("tgubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("bcode").Specific.VALUE = "";
                oForm.Items.Item("bname").Specific.VALUE = "";
                oForm.Items.Item("bnum").Specific.VALUE = "";
                oForm.Items.Item("yuncha").Specific.VALUE = 0;
                oForm.Items.Item("amt").Specific.VALUE = 0;
                oForm.Items.Item("gamt").Specific.VALUE = 0;

                oForm.Items.Item("amt").Enabled = true;
                ////Key set
                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;

                oForm.Items.Item("tyyyy").Enabled = false;
                oForm.Items.Item("tgubun").Enabled = false;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY411);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            PH_PY411_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY411_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY411_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY411_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        PH_PY411_DataFind();
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
                            PH_PY411_SAVE();
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
                            PH_PY411_Delete();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string Gubun = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        if (pVal.ItemUID == "gubun")
                        {
                            oForm.Items.Item("amt").Specific.VALUE = 0;
                            oForm.Items.Item("gamt").Specific.VALUE = 0;
                            Gubun = oForm.Items.Item("gubun").Specific.VALUE.Trim();

                            switch (Gubun)
                            {
                                case "61":
                                    oForm.Items.Item("tyyyy").Enabled = true;
                                    oForm.Items.Item("tgubun").Enabled = true;
                                    break;
                                default:
                                    oForm.Items.Item("tyyyy").Enabled = false;
                                    oForm.Items.Item("tgubun").Enabled = false;
                                    break;
                            }
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string FullName = string.Empty;
            string Gubun = string.Empty;
            string yyyy = string.Empty;
            string bcode = string.Empty;
            string seqn = string.Empty;
            double amt = 0;
            double gamt = 0;
            double samt = 0;
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
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

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

                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "FullName":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                FullName = oForm.Items.Item("FullName").Specific.VALUE.Trim();

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
                                sQry = sQry + " And U_status <> '5'";    // 퇴사자 제외
                                sQry = sQry + " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            //                        Case "gubun"
                            //
                            //
                            //                            Gubun = oForm.Items("gubun").Specific.VALUE
                            //
                            //                            Select Case Gubun
                            //                               Case "61"
                            //                                 oForm.Items("tyyyy").Enabled = True
                            //                                 oForm.Items("tgubun").Enabled = True
                            //                            End Select

                            case "bcode":
                                bcode = oForm.Items.Item("bcode").Specific.VALUE.Trim();
                                sQry = "Select Code,";
                                sQry = sQry + " CodeName = U_CodeNm ";
                                sQry = sQry + " From [@PS_HR200L]";
                                sQry = sQry + " WHERE Code = '78'";
                                sQry = sQry + " And U_Code = '" + bcode + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("bname").Specific.VALUE = oRecordSet.Fields.Item("CodeName").Value;
                                break;

                            case "amt":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();
                                seqn = oForm.Items.Item("seqn").Specific.VALUE.Trim();
                                amt = 0;
                                gamt = 0;
                                
                                Gubun = oForm.Items.Item("gubun").Specific.VALUE.Trim();

                                switch (Gubun)
                                {
                                    case "11":
                                    case "12":
                                    case "22":
                                        //11.근로자퇴직급여보장법, '12.과학기술인공제, 22.연금저축

                                        //총급여액계산해서 5,500 이하는 15% 아니면 12%
                                        sQry = "SELECT SUM(gwase) ";
                                        sQry = sQry + "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                        sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                        sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                        sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                        sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                        sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                        sQry = sQry + "      Union All ";
                                        sQry = sQry + "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                        sQry = sQry + "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                        sQry = sQry + "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                        sQry = sQry + "         And a.U_CLTCOD = b.U_CLTCOD ";
                                        sQry = sQry + "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                        sQry = sQry + "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                        sQry = sQry + "         And Isnull(b.U_PreCode,'') <> '' ";
                                        sQry = sQry + "     ) g";

                                        oRecordSet.DoQuery(sQry);
                                        samt = oRecordSet.Fields.Item(0).Value;  // 총급여액(과세대상)

                                        sQry = " Exec PH_PY411 '" + CLTCOD + "', '" + yyyy + "','" + MSTCOD + "','" + Gubun + "'," + oForm.Items.Item("amt").Specific.VALUE;
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        //                                   sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' And seqn <> '" & seqn & "' And gubun IN ('11','12','22') "
                                        //                                   oRecordSet.DoQuery sQry
                                        //                                   gamt = oRecordSet.Fields(0).VALUE

                                        //5500백기준
                                        if (samt <= 55000000)
                                        {
                                            amt = System.Math.Round(gamt * 0.15, 0); // 15%
                                            oForm.Items.Item("gamt").Specific.VALUE = amt;

                                        }
                                        else
                                        {
                                            amt = System.Math.Round(gamt * 0.12, 0); // 12%
                                            oForm.Items.Item("gamt").Specific.VALUE = amt;
                                        }
                                        oForm.Items.Item("amt").Specific.VALUE = gamt;
                                        break;

                                    case "21":
                                        //21.개인연금저축
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun = '21'";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) * 0.4, 0);

                                        if (gamt + amt > 720000)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 720000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.VALUE) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 0;
                                        }
                                        break;

                                    case "31":
                                    case "32":
                                    case "34":
                                        // 31.청약저축, 32.주택청약종합저축, 34.근로자주택마련저축
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun IN ('31','32','34') ";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) * 0.4, 0);

                                        if (gamt + amt > 960000)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 960000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.VALUE) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 0;
                                        }
                                        break;

                                    case "51":
                                        //51.장기집합투자증권저축  40% 240만원한도
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun = '51'";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) * 0.4, 0);

                                        if (gamt + amt > 2400000)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 2400000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.VALUE) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 0;
                                        }
                                        break;

                                    case "61":
                                        //61.중소기업창업투자조합출자 10%
                                        //2018년기준  2018년분은 개인투자조합,벤처기업에직접투자시 3천만원이하100%, 5천만원이하70%, 5천만원초과30%
                                        //            2016,2017년분은 개인투자조합,벤처기업에직접투자시 3천만원이하100%, 5천만원이하50%, 5천만원초과30%
                                        //종합(근로)소득금액의 50%한도
                                        //우리회사는해당사항이 없음 ..   있을시 계산필요

                                        //기본 10%만 계산
                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) * 0.1, 0);

                                        //종합(근로)소득금액의 50%한도 계산이 필요함........이상태에서는 어려움
                                        oForm.Items.Item("gamt").Specific.VALUE = amt;

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.VALUE) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.VALUE = 0;
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                string sQry = string.Empty;
                SAPbobsCOM.Recordset oRecordSet = null;
                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Param01, Param02, Param03, Param04 = string.Empty;

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);

                            Param01 = oDS_PH_PY411.Columns.Item("사업장").Cells.Item(pVal.Row).Value;
                            Param02 = oDS_PH_PY411.Columns.Item("년도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY411.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY411.Columns.Item("순번").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY411_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
                            oRecordSet.DoQuery(sQry);

                            if ((oRecordSet.RecordCount == 0))
                            {
                                oForm.Items.Item("seqn").Specific.VALUE = "";
                                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("tyyyy").Specific.VALUE = "";
                                oForm.Items.Item("tgubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("bcode").Specific.VALUE = "";
                                oForm.Items.Item("bname").Specific.VALUE = "";
                                oForm.Items.Item("bnum").Specific.VALUE = "";
                                oForm.Items.Item("yuncha").Specific.VALUE = 0;
                                oForm.Items.Item("amt").Specific.VALUE = 0;
                                oForm.Items.Item("gamt").Specific.VALUE = 0;

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");

                            }
                            else
                            {
                                oForm.DataSources.UserDataSources.Item("seqn").Value = oRecordSet.Fields.Item("seqn").Value;
                                oForm.Items.Item("gubun").Specific.Select(oRecordSet.Fields.Item("gubun").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("tyyyy").Value = oRecordSet.Fields.Item("tyyyy").Value;
                                oForm.Items.Item("tgubun").Specific.Select(oRecordSet.Fields.Item("tgubun").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("bcode").Value = oRecordSet.Fields.Item("bcode").Value;
                                oForm.DataSources.UserDataSources.Item("bname").Value = oRecordSet.Fields.Item("bname").Value;
                                oForm.DataSources.UserDataSources.Item("yuncha").Value = oRecordSet.Fields.Item("yuncha").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("amt").Value = oRecordSet.Fields.Item("amt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("gamt").Value = oRecordSet.Fields.Item("gamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("bnum").Value = oRecordSet.Fields.Item("bnum").Value;

                                // Key Disable
                                oForm.Items.Item("CLTCOD").Enabled = false;
                                oForm.Items.Item("Year").Enabled = false;
                                oForm.Items.Item("MSTCOD").Enabled = false;
                                
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY411_DataFind
        /// </summary>
        private void PH_PY411_DataFind()
        {
            int iRow = 0;
            string sQry = string.Empty;
            short ErrNum = 0;
            string CLTCOD, Year, MSTCOD = string.Empty;

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Strings.Trim(Year)))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                PH_PY411_FormItemEnabled();

                sQry = "EXEC PH_PY411_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY411.ExecuteQuery(sQry);
                iRow = oDS_PH_PY411.Rows.Count; //oForm.DataSources.DataTables.Item(0).Rows.Count;

                if(oDS_PH_PY411.IsEmpty) {
                    ErrNum = 3;
                    throw new Exception();
                }
                else {
                    PH_PY411_TitleSetting(iRow);
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1) {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년도가 없습니다. 확인바랍니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                } else if (ErrNum == 2) {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번이 없습니다. 확인바랍니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                } else if (ErrNum == 3) {
                    PSH_Globals.SBO_Application.StatusBar.SetText("등록된 자료가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY411_SAVE
        /// </summary>
        private void PH_PY411_SAVE()
        {
            // 데이타 저장
            int seqncom = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string saup, yyyy, sabun, seqn, Gubun, tyyyy, tgubun, bcode, bname, bnum = string.Empty;
            double yuncha, Amt, gamt = 0;

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                seqn = oForm.Items.Item("seqn").Specific.VALUE.ToString().Trim();
                Gubun = oForm.Items.Item("gubun").Specific.VALUE.ToString().Trim();
                tyyyy = oForm.Items.Item("tyyyy").Specific.VALUE.ToString().Trim();
                tgubun = oForm.Items.Item("tgubun").Specific.VALUE.ToString().Trim();
                bcode = oForm.Items.Item("bcode").Specific.VALUE.ToString().Trim();
                bname = oForm.Items.Item("bname").Specific.VALUE.ToString().Trim();
                bnum = oForm.Items.Item("bnum").Specific.VALUE.ToString().Trim();
                yuncha = Convert.ToDouble(oForm.Items.Item("yuncha").Specific.VALUE);
                Amt = Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE);
                gamt = Convert.ToDouble(oForm.Items.Item("gamt").Specific.VALUE);

                if (string.IsNullOrEmpty(Strings.Trim(yyyy)))
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(saup)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(sabun)))
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다..");
                    return;
                }
                
                if (string.IsNullOrEmpty(Strings.Trim(Gubun)) | string.IsNullOrEmpty(Strings.Trim(bcode)) | string.IsNullOrEmpty(Strings.Trim(bnum)) | Amt == 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select Count(*) From [p_seoybank] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신
                    sQry = "Update [p_seoybank] set ";
                    sQry = sQry + "gubun = '" + Gubun + "',";
                    sQry = sQry + "bcode = '" + bcode + "',";
                    sQry = sQry + "tyyyy = '" + tyyyy + "',";
                    sQry = sQry + "tgubun = '" + tgubun + "',";
                    sQry = sQry + "bname = '" + bname + "',";
                    sQry = sQry + "bnum = '" + bnum + "',";
                    sQry = sQry + "yuncha = " + yuncha + ",";
                    sQry = sQry + "amt = " + Amt + ",";
                    sQry = sQry + "gamt = " + gamt + "";
                    sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY411_DataFind();

                }
                else
                {
                    // 신규
                    //순번 계산
                    sQry = " Select Convert(int,Max(seqn)) From [p_seoybank] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    oRecordSet.DoQuery(sQry);
                    seqncom = Convert.ToInt16(oRecordSet.Fields.Item(0).Value);
                    seqncom = seqncom + 1;
                    seqn = seqncom.ToString().PadLeft(3, '0');
                    
                    //------------------------------------------------------

                    sQry = "INSERT INTO [p_seoybank]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "seqn,";
                    sQry = sQry + "gubun,";
                    sQry = sQry + "tyyyy,";
                    sQry = sQry + "tgubun,";
                    sQry = sQry + "bcode,";
                    sQry = sQry + "bname,";
                    sQry = sQry + "bnum,";
                    sQry = sQry + "yuncha,";
                    sQry = sQry + "amt,";
                    sQry = sQry + "gamt";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + saup + "',";
                    sQry = sQry + "'" + yyyy + "',";
                    sQry = sQry + "'" + sabun + "',";
                    sQry = sQry + "'" + seqn + "',";
                    sQry = sQry + "'" + Gubun + "',";
                    sQry = sQry + "'" + tyyyy + "',";
                    sQry = sQry + "'" + tgubun + "',";
                    sQry = sQry + "'" + bcode + "',";
                    sQry = sQry + "'" + bname + "',";
                    sQry = sQry + "'" + bnum + "',";
                    sQry = sQry + yuncha + ",";
                    sQry = sQry + Amt + ",";
                    sQry = sQry + gamt + "";
                    sQry = sQry + " ) ";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY411_DataFind();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY411_Delete
        /// </summary>
        private void PH_PY411_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string saup, yyyy, sabun, seqn = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                yyyy = oForm.Items.Item("Year").Specific.VALUE;
                sabun = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
                seqn = oForm.Items.Item("seqn").Specific.VALUE;

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY411.Rows.Count > 0)
                    {
                        sQry = "Delete From [p_seoybank] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY411_DataFind();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY411_TitleSetting
        /// </summary>
        private void PH_PY411_TitleSetting(int iRow)
        {
            int i = 0;
            string[] COLNAM = new string[14];

            try
            {
                //
                COLNAM[0] = "순번";
                COLNAM[1] = "구분코드";
                COLNAM[2] = "구분명";
                COLNAM[3] = "금융기관코드";
                COLNAM[4] = "금융기관명";
                COLNAM[5] = "계좌번호";
                COLNAM[6] = "납입년차";
                COLNAM[7] = "납입금액";
                COLNAM[8] = "공제금액";
                COLNAM[9] = "사번";
                COLNAM[10] = "년도";
                COLNAM[11] = "사업장";
                COLNAM[12] = "투자년도";
                COLNAM[13] = "투자구분";
                

                for (i = 0; i <= Information.UBound(COLNAM); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i <= Information.UBound(COLNAM))
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                }
                oGrid1.Columns.Item(6).RightJustified = true;
                oGrid1.Columns.Item(7).RightJustified = true;
                oGrid1.Columns.Item(8).RightJustified = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

    }
}
