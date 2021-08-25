using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연금저축등소득공제등록
    /// </summary>
    internal class PH_PY411 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY411;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
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

                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                }
                oForm.Items.Item("seqn").Specific.Value = "";
                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("tyyyy").Specific.Value = "";
                oForm.Items.Item("tgubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("bcode").Specific.Value = "";
                oForm.Items.Item("bname").Specific.Value = "";
                oForm.Items.Item("bnum").Specific.Value = "";
                oForm.Items.Item("yuncha").Specific.Value = 0;
                oForm.Items.Item("amt").Specific.Value = 0;
                oForm.Items.Item("gamt").Specific.Value = 0;

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY411);
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
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    break;
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
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        yyyy = oForm.Items.Item("Year").Specific.Value;
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

                        yyyy = oForm.Items.Item("Year").Specific.Value;
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
            string Gubun = string.Empty;

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
                            oForm.Items.Item("amt").Specific.Value = 0;
                            oForm.Items.Item("gamt").Specific.Value = 0;
                            Gubun = oForm.Items.Item("gubun").Specific.Value.Trim();
                            oForm.ActiveItem = "bcode";

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
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();

                                sQry  = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "FullName":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                FullName = oForm.Items.Item("FullName").Specific.Value.Trim();

                                sQry  = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += " And U_status <> '5'";    // 퇴사자 제외
                                sQry += " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "bcode":
                                bcode = oForm.Items.Item("bcode").Specific.Value.Trim();
                                sQry  = "Select Code,";
                                sQry += " CodeName = U_CodeNm ";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '78'";
                                sQry += " And U_Code = '" + bcode + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("bname").Specific.Value = oRecordSet.Fields.Item("CodeName").Value;
                                break;

                            case "amt":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                                seqn = oForm.Items.Item("seqn").Specific.Value.ToString().Trim();
                                amt = 0;
                                gamt = 0;
                                
                                Gubun = oForm.Items.Item("gubun").Specific.Value.ToString().Trim();

                                switch (Gubun)
                                {
                                    case "11":
                                    case "12":
                                    case "22":
                                        //11.근로자퇴직급여보장법, '12.과학기술인공제, 22.연금저축

                                        // 총급여액계산해서 5,500 이하는 15% 아니면 12%
                                        // 2020년 50세이상 공제한도 확대(3년한시 2022.12.31까지)
                                        sQry = " Exec PH_PY411 '" + CLTCOD + "', '" + yyyy + "','" + MSTCOD + "','" + Gubun + "'," + oForm.Items.Item("amt").Specific.Value;
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;  // 불입금액
                                        samt = oRecordSet.Fields.Item(1).Value;  // 총급여액(과세대상)

                                        //5500백기준
                                        if (samt <= 55000000)
                                        {
                                            amt = System.Math.Round(gamt * 0.15, 0); // 15%
                                            oForm.Items.Item("gamt").Specific.Value = amt;
                                        }
                                        else
                                        {
                                            amt = System.Math.Round(gamt * 0.12, 0); // 12%
                                            oForm.Items.Item("gamt").Specific.Value = amt;
                                        }
                                        oForm.Items.Item("amt").Specific.Value = gamt;
                                        oForm.Items.Item("Age").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim() + " 세" ;
                                        break;

                                    case "21":
                                        //21.개인연금저축
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun = '21'";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.Value) * 0.4, 0);

                                        if (gamt + amt > 720000)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 720000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 0;
                                        }
                                        break;

                                    case "31":
                                    case "32":
                                    case "34":
                                        // 31.청약저축, 32.주택청약종합저축, 34.근로자주택마련저축
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun IN ('31','32','34') ";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.Value) * 0.4, 0);

                                        if (gamt + amt > 960000)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 960000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 0;
                                        }
                                        break;

                                    case "51":
                                        //51.장기집합투자증권저축  40% 240만원한도
                                        sQry = " Select sum(gamt) From [p_seoybank] Where saup = '" + CLTCOD + "' And yyyy = '" + yyyy + "' And sabun = '" + MSTCOD + "' And seqn <> '" + seqn + "' And gubun = '51'";
                                        oRecordSet.DoQuery(sQry);
                                        gamt = oRecordSet.Fields.Item(0).Value;

                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.Value) * 0.4, 0);

                                        if (gamt + amt > 2400000)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 2400000 - gamt;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = amt;
                                        }

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 0;
                                        }
                                        break;

                                    case "61":
                                        //61.중소기업창업투자조합출자 10%
                                        //2018년기준  2018년분은 개인투자조합,벤처기업에직접투자시 3천만원이하100%, 5천만원이하70%, 5천만원초과30%
                                        //            2016,2017년분은 개인투자조합,벤처기업에직접투자시 3천만원이하100%, 5천만원이하50%, 5천만원초과30%
                                        //종합(근로)소득금액의 50%한도
                                        //우리회사는해당사항이 없음 ..   있을시 계산필요

                                        //기본 10%만 계산
                                        amt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("amt").Specific.Value) * 0.1, 0);

                                        //종합(근로)소득금액의 50%한도 계산이 필요함........이상태에서는 어려움
                                        oForm.Items.Item("gamt").Specific.Value = amt;

                                        if (Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value) < 0)
                                        {
                                            oForm.Items.Item("gamt").Specific.Value = 0;
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
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            string Param04 = string.Empty;
            
            try
            {
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
                                oForm.Items.Item("seqn").Specific.Value = "";
                                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("tyyyy").Specific.Value = "";
                                oForm.Items.Item("tgubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("bcode").Specific.Value = "";
                                oForm.Items.Item("bname").Specific.Value = "";
                                oForm.Items.Item("bnum").Specific.Value = "";
                                oForm.Items.Item("yuncha").Specific.Value = 0;
                                oForm.Items.Item("amt").Specific.Value = 0;
                                oForm.Items.Item("gamt").Specific.Value = 0;

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
                            oForm.ActiveItem = "gubun";
                            oForm.Items.Item("amt").Enabled = false;   // 2020년 수정 못하게 막음  삭제후 등록해야 됨
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string CLTCOD = string.Empty;
            string Year = string.Empty;
            string MSTCOD = string.Empty;

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Year))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(MSTCOD))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                PH_PY411_FormItemEnabled();

                sQry = "EXEC PH_PY411_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY411.ExecuteQuery(sQry);
                iRow = oDS_PH_PY411.Rows.Count; //oForm.DataSources.DataTables.Item(0).Rows.Count;

                PH_PY411_TitleSetting(iRow);
                oForm.ActiveItem = "gubun";
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1) {
                    PSH_Globals.SBO_Application.StatusBar.SetText("년도가 없습니다. 확인바랍니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                } else if (ErrNum == 2) {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번이 없습니다. 확인바랍니다..", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
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

            string saup = string.Empty;
            string yyyy = string.Empty;
            string sabun = string.Empty;
            string seqn = string.Empty;
            string Gubun = string.Empty;
            string tyyyy = string.Empty;
            string tgubun = string.Empty;
            string bcode = string.Empty;
            string bname = string.Empty;
            string bnum = string.Empty;
            double yuncha = 0;
            double Amt = 0;
            double gamt = 0;

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                seqn = oForm.Items.Item("seqn").Specific.Value.ToString().Trim();
                Gubun = oForm.Items.Item("gubun").Specific.Value.ToString().Trim();
                tyyyy = oForm.Items.Item("tyyyy").Specific.Value.ToString().Trim();
                tgubun = oForm.Items.Item("tgubun").Specific.Value.ToString().Trim();
                bcode = oForm.Items.Item("bcode").Specific.Value.ToString().Trim();
                bname = oForm.Items.Item("bname").Specific.Value.ToString().Trim();
                bnum = oForm.Items.Item("bnum").Specific.Value.ToString().Trim();
                yuncha = Convert.ToDouble(oForm.Items.Item("yuncha").Specific.Value);
                Amt = Convert.ToDouble(oForm.Items.Item("amt").Specific.Value);
                gamt = Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value);

                if (string.IsNullOrEmpty(yyyy))
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(saup))
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(sabun))
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다..");
                    return;
                }
                
                if (string.IsNullOrEmpty(Gubun) || string.IsNullOrEmpty(bcode) || string.IsNullOrEmpty(bnum) || Amt == 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select Count(*) From [p_seoybank] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신

                    PSH_Globals.SBO_Application.StatusBar.SetText("수정할수 없습니다. 자료를 삭제후 수정하여 재등록 하세요...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    //sQry  = "Update [p_seoybank] set ";
                    //sQry += "gubun = '" + Gubun + "',";
                    //sQry += "bcode = '" + bcode + "',";
                    //sQry += "tyyyy = '" + tyyyy + "',";
                    //sQry += "tgubun = '" + tgubun + "',";
                    //sQry += "bname = '" + bname + "',";
                    //sQry += "bnum = '" + bnum + "',";
                    //sQry += "yuncha = " + yuncha + ",";
                    //sQry += "amt = " + Amt + ",";
                    //sQry += "gamt = " + gamt + "";
                    //sQry += " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";

                    //oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    //PSH_Globals.SBO_Application.StatusBar.SetText("자료가 수정 되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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

                    sQry  = "INSERT INTO [p_seoybank]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "seqn,";
                    sQry += "gubun,";
                    sQry += "tyyyy,";
                    sQry += "tgubun,";
                    sQry += "bcode,";
                    sQry += "bname,";
                    sQry += "bnum,";
                    sQry += "yuncha,";
                    sQry += "amt,";
                    sQry += "gamt";
                    sQry += " ) ";
                    sQry += "VALUES(";

                    sQry += "'" + saup + "',";
                    sQry += "'" + yyyy + "',";
                    sQry += "'" + sabun + "',";
                    sQry += "'" + seqn + "',";
                    sQry += "'" + Gubun + "',";
                    sQry += "'" + tyyyy + "',";
                    sQry += "'" + tgubun + "',";
                    sQry += "'" + bcode + "',";
                    sQry += "'" + bname + "',";
                    sQry += "'" + bnum + "',";
                    sQry += yuncha + ",";
                    sQry += Amt + ",";
                    sQry += gamt + "";
                    sQry += " ) ";

                    oRecordSet.DoQuery(sQry);
                    PSH_Globals.SBO_Application.StatusBar.SetText("자료가 저장 되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY411_DataFind();
                }
                oForm.Items.Item("Age").Specific.Value = "";
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                oForm.Items.Item("Age").Specific.Value = "";
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
            string saup = string.Empty;
            string yyyy = string.Empty;
            string sabun = string.Empty;
            string seqn = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                seqn = oForm.Items.Item("seqn").Specific.Value;

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    if (oDS_PH_PY411.Rows.Count > 0)
                    {
                        sQry = "Delete From [p_seoybank] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And seqn = '" + seqn + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PSH_Globals.SBO_Application.StatusBar.SetText("자료가 삭제 되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        PH_PY411_DataFind();
                    }
                }
                oGrid1.AutoResizeColumns();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i <= COLNAM.Length )
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                }
                oGrid1.Columns.Item(6).RightJustified = true;
                oGrid1.Columns.Item(7).RightJustified = true;
                oGrid1.Columns.Item(8).RightJustified = true;
                oGrid1.AutoResizeColumns();
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

    }
}
