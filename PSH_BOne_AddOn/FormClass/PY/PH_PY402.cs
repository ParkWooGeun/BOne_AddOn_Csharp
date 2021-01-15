using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정산기초등록
    /// </summary>
    internal class PH_PY402 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        // 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.Matrix oMat01;
        public SAPbouiCOM.DataTable oDS_PH_PY402A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY402L;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY402.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY402_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY402");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY402_CreateItems();
                PH_PY402_FormItemEnabled();
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
        private void PH_PY402_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY402L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oGrid1 = oForm.Items.Item("Grid01").Specific;

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oForm.DataSources.DataTables.Add("PH_PY402");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY402");
                oDS_PH_PY402A = oForm.DataSources.DataTables.Item("PH_PY402");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("공제구분코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("공제구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("공제대상코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("공제대상", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("관계코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("관계", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("주민번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("금액(국세청)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("금액(국세청외)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("전통시장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("대중교통", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("도서공연", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY402").Columns.Add("합계금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                //성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 공제구분
                oForm.DataSources.UserDataSources.Add("div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("div").Specific.DataBind.SetBound(true, "", "div");

                // 공제구분명
                oForm.DataSources.UserDataSources.Add("divnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("divnm").Specific.DataBind.SetBound(true, "", "divnm");

                // 공제대상
                oForm.DataSources.UserDataSources.Add("target", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("target").Specific.DataBind.SetBound(true, "", "target");

                // 공제대상명
                oForm.DataSources.UserDataSources.Add("targetnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("targetnm").Specific.DataBind.SetBound(true, "", "targetnm");

                // 관계
                oForm.DataSources.UserDataSources.Add("relate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("relate").Specific, "Y");

                // 성명
                oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

                // 주민번호
                oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

                // 생년월일
                oForm.DataSources.UserDataSources.Add("birthymd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("birthymd").Specific.DataBind.SetBound(true, "", "birthymd");

                // 주소
                oForm.DataSources.UserDataSources.Add("addr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("addr").Specific.DataBind.SetBound(true, "", "addr");
                oForm.Items.Item("addr").Enabled = false;

                // 공제금액(국세청)
                oForm.DataSources.UserDataSources.Add("ntsamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntsamt").Specific.DataBind.SetBound(true, "", "ntsamt");

                // 공제금액(국세청외)
                oForm.DataSources.UserDataSources.Add("amt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("amt").Specific.DataBind.SetBound(true, "", "amt");

                // 한도금액
                oForm.DataSources.UserDataSources.Add("handoamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("handoamt").Specific.DataBind.SetBound(true, "", "handoamt");

                // 일반금액_3월
                oForm.DataSources.UserDataSources.Add("ntsamt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntsamt3").Specific.DataBind.SetBound(true, "", "ntsamt3");

                // 전통시장_3월
                oForm.DataSources.UserDataSources.Add("mart3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("mart3").Specific.DataBind.SetBound(true, "", "mart3");

                // 대중교통_3월
                oForm.DataSources.UserDataSources.Add("trans3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("trans3").Specific.DataBind.SetBound(true, "", "trans3");

                // 도서공연_3월
                oForm.DataSources.UserDataSources.Add("bookpms3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bookpms3").Specific.DataBind.SetBound(true, "", "bookpms3");

                // 일반금액_4-7월
                oForm.DataSources.UserDataSources.Add("ntsamt47", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntsamt47").Specific.DataBind.SetBound(true, "", "ntsamt47");

                // 전통시장_4-7월
                oForm.DataSources.UserDataSources.Add("mart47", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("mart47").Specific.DataBind.SetBound(true, "", "mart47");

                // 대중교통_4-7월
                oForm.DataSources.UserDataSources.Add("trans47", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("trans47").Specific.DataBind.SetBound(true, "", "trans47");

                // 도서공연_4-7월
                oForm.DataSources.UserDataSources.Add("bookpms47", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bookpms47").Specific.DataBind.SetBound(true, "", "bookpms47");

                // 일반금액_그외
                oForm.DataSources.UserDataSources.Add("ntsamt24", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntsamt24").Specific.DataBind.SetBound(true, "", "ntsamt24");

                // 전통시장_그외
                oForm.DataSources.UserDataSources.Add("mart24", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("mart24").Specific.DataBind.SetBound(true, "", "mart24");

                // 대중교통_그외
                oForm.DataSources.UserDataSources.Add("trans24", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("trans24").Specific.DataBind.SetBound(true, "", "trans24");

                // 도서공연_그외
                oForm.DataSources.UserDataSources.Add("bookpms", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("bookpms").Specific.DataBind.SetBound(true, "", "bookpms");

                //장애인코드
                oForm.DataSources.UserDataSources.Add("hdcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("hdcode").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("hdcode").Specific.ValidValues.Add("1", "장애인복지법에 따른 장애인");
                oForm.Items.Item("hdcode").Specific.ValidValues.Add("2", "국가유공자등 예우및지원에 관한 법률에 따른 상이자 및 이와 유사한자로서 근로능력이없는자");
                oForm.Items.Item("hdcode").Specific.ValidValues.Add("3", "그 밖에 항시 치료를 요하는 중증환자");
                oForm.Items.Item("hdcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY402_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY402_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // 문서추가
                oForm.EnableMenu("1282", true);


                if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE)))
                {
                    oForm.Items.Item("Year").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);
                }

                if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE)))
                {
                    oForm.Items.Item("MSTCOD").Specific.VALUE = "";
                    oForm.Items.Item("FullName").Specific.VALUE = "";
                    oForm.Items.Item("TeamName").Specific.VALUE = "";
                    oForm.Items.Item("RspName").Specific.VALUE = "";
                    oForm.Items.Item("ClsName").Specific.VALUE = "";
                }

                oForm.DataSources.UserDataSources.Item("div").Value = "";
                oForm.DataSources.UserDataSources.Item("divnm").Value = "";
                oForm.DataSources.UserDataSources.Item("target").Value = "";
                oForm.DataSources.UserDataSources.Item("targetnm").Value = "";
                oForm.Items.Item("relate").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("hdcode").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                oForm.DataSources.UserDataSources.Item("birthymd").Value = "";
                oForm.DataSources.UserDataSources.Item("addr").Value = "";
                oForm.DataSources.UserDataSources.Item("ntsamt").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("handoamt").Value = Convert.ToString(0);

                oForm.Items.Item("ntsamt").Enabled = true;
                oForm.Items.Item("ntsamt24").Enabled = false;
                oForm.Items.Item("mart24").Enabled = false;
                oForm.Items.Item("trans24").Enabled = false;
                oForm.Items.Item("bookpms").Enabled = false;

                oForm.Items.Item("ntsamt3").Enabled = false;
                oForm.Items.Item("mart3").Enabled = false;
                oForm.Items.Item("trans3").Enabled = false;
                oForm.Items.Item("bookpms3").Enabled = false;

                oForm.Items.Item("ntsamt47").Enabled = false;
                oForm.Items.Item("mart47").Enabled = false;
                oForm.Items.Item("trans47").Enabled = false;
                oForm.Items.Item("bookpms47").Enabled = false;

                oForm.DataSources.UserDataSources.Item("ntsamt24").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("mart24").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("trans24").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("bookpms").Value = Convert.ToString(0);

                oForm.DataSources.UserDataSources.Item("ntsamt3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("mart3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("trans3").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("bookpms3").Value = Convert.ToString(0);

                oForm.DataSources.UserDataSources.Item("ntsamt47").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("mart47").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("trans47").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("bookpms47").Value = Convert.ToString(0);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY402_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY402L);
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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                            PH_PY402_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY402_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY402_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY402_FormItemEnabled();
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
                        PH_PY402_DataFind();
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
                            PH_PY402_SAVE();
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
                            PH_PY402_Delete();
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "MSTCOD")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                                BubbleEvent = false;
                            }
                        }

                        if (pVal.ItemUID == "div")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("div").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "target")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("target").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                                BubbleEvent = false;
                            }
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        oForm.Items.Item("kname").Specific.VALUE = oMat01.Columns.Item("kname").Cells.Item(pVal.Row).Specific.VALUE;
                        oForm.Items.Item("juminno").Specific.VALUE = oMat01.Columns.Item("juminno").Cells.Item(pVal.Row).Specific.VALUE;
                        oForm.Items.Item("birthymd").Specific.VALUE = oMat01.Columns.Item("birthymd").Cells.Item(pVal.Row).Specific.VALUE;
                        oForm.Items.Item("addr").Specific.VALUE = oMat01.Columns.Item("addr").Cells.Item(pVal.Row).Specific.VALUE;
                    }
                    // 신용카드(520,540,550)일때
                    if (oForm.Items.Item("target").Specific.VALUE == "520" | oForm.Items.Item("target").Specific.VALUE == "540" | oForm.Items.Item("target").Specific.VALUE == "550")
                    {
                        oForm.Items.Item("ntsamt24").Click(SAPbouiCOM.BoCellClickType.ct_Regular);  // 포커싱을 일반금액으로..
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            int i = 0;
            string sQry = string.Empty;
            string MSTCOD = string.Empty;
            string relate = string.Empty;
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

                        if (pVal.ItemUID == "relate")
                        {
                            oMat01.Clear();
                            oDS_PH_PY402L.Clear();

                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                            relate = oForm.Items.Item("relate").Specific.VALUE;

                            sQry = "EXEC [PH_PY402_03] '" + MSTCOD + "', '" + relate + "'";

                            oRecordSet.DoQuery(sQry);

                            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                            {
                                if (i + 1 > oDS_PH_PY402L.Size)
                                {
                                    oDS_PH_PY402L.InsertRecord((i));
                                }

                                oMat01.AddRow();
                                oDS_PH_PY402L.Offset = i;

                                oDS_PH_PY402L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                oDS_PH_PY402L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet.Fields.Item("kname").Value));
                                oDS_PH_PY402L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet.Fields.Item("juminno").Value));
                                oDS_PH_PY402L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet.Fields.Item("birthymd").Value));
                                oDS_PH_PY402L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet.Fields.Item("relatenm").Value));
                                oDS_PH_PY402L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet.Fields.Item("addr").Value));
                                oRecordSet.MoveNext();
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();

                            if ((oRecordSet.RecordCount == 0))
                            {
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.DataSources.UserDataSources.Item("birthymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("addr").Value = "";
                            }

                            if ((oRecordSet.RecordCount == 1))
                            {
                                oForm.Items.Item("kname").Specific.VALUE = oMat01.Columns.Item("kname").Cells.Item(1).Specific.VALUE;
                                oForm.Items.Item("juminno").Specific.VALUE = oMat01.Columns.Item("juminno").Cells.Item(1).Specific.VALUE;
                                oForm.Items.Item("birthymd").Specific.VALUE = oMat01.Columns.Item("birthymd").Cells.Item(1).Specific.VALUE;
                                oForm.Items.Item("addr").Specific.VALUE = oMat01.Columns.Item("addr").Cells.Item(1).Specific.VALUE;
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string Div = string.Empty;
            string target = string.Empty;
            string YEAR_Renamed = string.Empty;
            Double bookAmt = 0;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

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

                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "FullName":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                FullName = oForm.Items.Item("FullName").Specific.VALUE;

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
                                sQry += " And U_status <> '5'"; // 퇴사자 제외
                                sQry += " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "div":
                                Div = oForm.Items.Item("div").Specific.VALUE;

                                sQry  = "Select CodeNm = U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '70'";
                                sQry += " And U_Code = '" + Div + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("divnm").Specific.VALUE = oRecordSet.Fields.Item("CodeNm").Value;
                                break;

                            case "target":
                                target = oForm.Items.Item("target").Specific.VALUE;

                                sQry  = "Select CodeNm = U_CodeNm, handoamt = Isnull(U_Num1,0)";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '71'";
                                sQry += " And U_Code = '" + target + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("targetnm").Specific.VALUE = oRecordSet.Fields.Item("CodeNm").Value;
                                oForm.Items.Item("handoamt").Specific.VALUE = Convert.ToString(oRecordSet.Fields.Item("handoamt").Value);

                                if (target == "520" || target == "540" || target == "550")
                                {
                                    // 신용카드(520,540,550)일때
                                    oForm.Items.Item("ntsamt").Enabled = false;
                                    oForm.Items.Item("amt").Enabled = false;      // 2020년부터 국세청외(기타)는 없애기로 모두 국세청으로 등록
                                    oForm.Items.Item("ntsamt3").Click(SAPbouiCOM.BoCellClickType.ct_Regular);  // 포커싱을 일반금액으로..

                                    oForm.Items.Item("ntsamt24").Enabled = true;
                                    oForm.Items.Item("mart24").Enabled = true;
                                    oForm.Items.Item("trans24").Enabled = true;
                                    oForm.Items.Item("bookpms").Enabled = true;

                                    oForm.Items.Item("ntsamt3").Enabled = true;
                                    oForm.Items.Item("mart3").Enabled = true;
                                    oForm.Items.Item("trans3").Enabled = true;
                                    oForm.Items.Item("bookpms3").Enabled = true;

                                    oForm.Items.Item("ntsamt47").Enabled = true;
                                    oForm.Items.Item("mart47").Enabled = true;
                                    oForm.Items.Item("trans47").Enabled = true;
                                    oForm.Items.Item("bookpms47").Enabled = true;

                                    oForm.Items.Item("ntsamt24").Specific.VALUE = 0;
                                    oForm.Items.Item("mart24").Specific.VALUE = 0;
                                    oForm.Items.Item("trans24").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms").Specific.VALUE = 0;

                                    oForm.Items.Item("ntsamt3").Specific.VALUE = 0;
                                    oForm.Items.Item("mart3").Specific.VALUE = 0;
                                    oForm.Items.Item("trans3").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms3").Specific.VALUE = 0;

                                    oForm.Items.Item("ntsamt47").Specific.VALUE = 0;
                                    oForm.Items.Item("mart47").Specific.VALUE = 0;
                                    oForm.Items.Item("trans47").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms47").Specific.VALUE = 0;
                                }
                                else
                                {
                                    oForm.Items.Item("ntsamt").Enabled = true;
                                    oForm.Items.Item("amt").Enabled = true;

                                    oForm.Items.Item("ntsamt24").Enabled = false;
                                    oForm.Items.Item("mart24").Enabled = false;
                                    oForm.Items.Item("trans24").Enabled = false;
                                    oForm.Items.Item("bookpms").Enabled = false;

                                    oForm.Items.Item("ntsamt3").Enabled = false;
                                    oForm.Items.Item("mart3").Enabled = false;
                                    oForm.Items.Item("trans3").Enabled = false;
                                    oForm.Items.Item("bookpms3").Enabled = false;

                                    oForm.Items.Item("ntsamt47").Enabled = false;
                                    oForm.Items.Item("mart47").Enabled = false;
                                    oForm.Items.Item("trans47").Enabled = false;
                                    oForm.Items.Item("bookpms47").Enabled = false;

                                    oForm.Items.Item("ntsamt24").Specific.VALUE = 0;
                                    oForm.Items.Item("mart24").Specific.VALUE = 0;
                                    oForm.Items.Item("trans24").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms").Specific.VALUE = 0;

                                    oForm.Items.Item("ntsamt3").Specific.VALUE = 0;
                                    oForm.Items.Item("mart3").Specific.VALUE = 0;
                                    oForm.Items.Item("trans3").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms3").Specific.VALUE = 0;

                                    oForm.Items.Item("ntsamt47").Specific.VALUE = 0;
                                    oForm.Items.Item("mart47").Specific.VALUE = 0;
                                    oForm.Items.Item("trans47").Specific.VALUE = 0;
                                    oForm.Items.Item("bookpms47").Specific.VALUE = 0;
                                }

                                switch (target)
                                {
                                    case "110":
                                        // 본인
                                        oForm.Items.Item("relate").Specific.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("amt").Specific.VALUE = oForm.Items.Item("handoamt").Specific.VALUE;
                                        break;
                                    case "120":
                                        // 배우자
                                        oForm.Items.Item("relate").Specific.Select("02", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("amt").Specific.VALUE = oForm.Items.Item("handoamt").Specific.VALUE;
                                        break;

                                    case "130":
                                        // 부양가족
                                        oForm.Items.Item("amt").Specific.VALUE = oForm.Items.Item("handoamt").Specific.VALUE;
                                        break;

                                    default:
                                        oForm.Items.Item("relate").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oMat01.Clear();

                                        oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                        oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                        oForm.DataSources.UserDataSources.Item("birthymd").Value = "";
                                        oForm.DataSources.UserDataSources.Item("addr").Value = "";

                                        if (oForm.Items.Item("div").Specific.VALUE == "20")
                                        {

                                            if (Convert.ToDouble(oForm.DataSources.UserDataSources.Item("handoamt").Value) > 0)
                                            {
                                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = Convert.ToString(0);
                                                oForm.DataSources.UserDataSources.Item("amt").Value = oForm.DataSources.UserDataSources.Item("handoamt").Value;
                                            }
                                            else
                                            {
                                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = Convert.ToString(0);
                                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                            }
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(oForm.DataSources.UserDataSources.Item("handoamt").Value) > 0)
                                            {
                                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = oForm.DataSources.UserDataSources.Item("handoamt").Value;
                                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                            }
                                            else
                                            {
                                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = Convert.ToString(0);
                                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                            }
                                        }
                                        break;

                                }
                                break;

                            case "juminno":
                                // 주민번호
                                // 주민번호입력시 생년월일 생성
                                if (Strings.Len(Strings.Trim(oForm.Items.Item("juminno").Specific.VALUE)) != 13)
                                {
                                    oForm.Items.Item("birthymd").Specific.VALUE = "";
                                    PSH_Globals.SBO_Application.MessageBox("주민번호자릿수가 틀립니다. 확인하세요.");
                                }
                                else
                                {
                                    if (Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "1" | Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "2")
                                    {
                                        oForm.Items.Item("birthymd").Specific.VALUE = "19" + Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 1, 6);
                                    }
                                    else if (Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "3" | Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "4")
                                    {
                                        oForm.Items.Item("birthymd").Specific.VALUE = "20" + Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 1, 6);
                                    }
                                    else if (Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "5" | Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "6")
                                    {
                                        oForm.Items.Item("birthymd").Specific.VALUE = "19" + Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 1, 6);
                                    }
                                    else if (Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "7" | Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 7, 1) == "8")
                                    {
                                        oForm.Items.Item("birthymd").Specific.VALUE = "20" + Strings.Mid(oForm.Items.Item("juminno").Specific.VALUE, 1, 6);
                                    }
                                }
                                break;

                            case "ntsamt":
                                if (Convert.ToDouble(oForm.Items.Item("handoamt").Specific.Value) > 0 && ( oForm.Items.Item("target").Specific.VALUE == "633" && oForm.Items.Item("relate").Specific.VALUE != "01") )
                                // 대학교육비 본인은 한도 없슴
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("ntsamt").Specific.VALUE) + Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) > Convert.ToDouble(oForm.Items.Item("handoamt").Specific.VALUE))
                                    {
                                        oForm.Items.Item("ntsamt").Specific.VALUE = 0;
                                        PSH_Globals.SBO_Application.MessageBox("한도금액보다 초과됩니다. 확인하세요");
                                    }
                                }
                                break;
                            
                            case "amt":
                                if (Convert.ToDouble(oForm.Items.Item("handoamt").Specific.Value) > 0 && (oForm.Items.Item("target").Specific.VALUE == "633" && oForm.Items.Item("relate").Specific.VALUE != "01"))
                                // 대학교육비 본인은 한도 없슴
                                {
                                    if (Convert.ToDouble(oForm.Items.Item("ntsamt").Specific.VALUE) + Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE) > Convert.ToDouble(oForm.Items.Item("handoamt").Specific.VALUE))
                                    {
                                        oForm.Items.Item("amt").Specific.VALUE = 0;
                                        PSH_Globals.SBO_Application.MessageBox("한도금액보다 초과됩니다. 확인하세요");
                                    }
                                }
                                break;
                            case "ntsamt3":
                                oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt3").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt47").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt24").Specific.VALUE);
                                break;
                            case "ntsamt47":
                                oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt3").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt47").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt24").Specific.VALUE);
                                break;
                            case "ntsamt24":
                                oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt3").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt47").Specific.VALUE) + Conversion.Val(oForm.Items.Item("ntsamt24").Specific.VALUE);
                                break;

                            //2018부터 도서공연사용분 총급여 7천만원 CHECK
                            case "bookpms3":
                                //도서공연사용분
                                //총급여액계산해서 7천만원이하는 0
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                YEAR_Renamed = oForm.Items.Item("Year").Specific.VALUE;
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                                bookAmt = 0;

                                sQry  = "SELECT SUM(gwase) ";
                                sQry += "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "      Union All ";
                                sQry += "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry += "      Union All";
                                sQry += "      SELECT gwase   = SUM( isnull(a.payrtot1 ,0) + isnull(a.payrtot2,0) + isnull(a.bnstot1,0) + isnull(a.bnstot2,0) )";
                                sQry += "        FROM p_sbservcomp a";
                                sQry += "       WHERE a.saup = '" + CLTCOD + "' ";
                                sQry += "         And a.yyyy   =  '" + YEAR_Renamed + "'";
                                sQry += "         And a.sabun  = '" + MSTCOD + "' ";
                                sQry += "     ) g";

                                oRecordSet.DoQuery(sQry);
                                bookAmt = oRecordSet.Fields.Item(0).Value;
                                //총급여액(과세대상)
                                //7천기준
                                if (bookAmt > 70000000)
                                {
                                    oForm.Items.Item("ntsamt3").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt3").Specific.VALUE) + Conversion.Val(oForm.Items.Item("bookpms3").Specific.VALUE);
                                    oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt3").Specific.VALUE);
                                    oForm.Items.Item("bookpms3").Specific.VALUE = 0;
                                    PSH_Globals.SBO_Application.MessageBox("총급여 7천만원 초과자입니다. 일반금액에 합산하고 도서공연비는 0처리 합니다.");
                                }
                                break;
                            case "bookpms47":
                                //도서공연사용분
                                //총급여액계산해서 7천만원이하는 0
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                YEAR_Renamed = oForm.Items.Item("Year").Specific.VALUE;
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                                bookAmt = 0;

                                sQry  = "SELECT SUM(gwase) ";
                                sQry += "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "      Union All ";
                                sQry += "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry += "      Union All";
                                sQry += "      SELECT gwase   = SUM( isnull(a.payrtot1 ,0) + isnull(a.payrtot2,0) + isnull(a.bnstot1,0) + isnull(a.bnstot2,0) )";
                                sQry += "        FROM p_sbservcomp a";
                                sQry += "       WHERE a.saup = '" + CLTCOD + "' ";
                                sQry += "         And a.yyyy   =  '" + YEAR_Renamed + "'";
                                sQry += "         And a.sabun  = '" + MSTCOD + "' ";
                                sQry += "     ) g";

                                oRecordSet.DoQuery(sQry);
                                bookAmt = oRecordSet.Fields.Item(0).Value;
                                //총급여액(과세대상)
                                //7천기준
                                if (bookAmt > 70000000)
                                {
                                    oForm.Items.Item("ntsamt47").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt47").Specific.VALUE) + Conversion.Val(oForm.Items.Item("bookpms47").Specific.VALUE);
                                    oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt47").Specific.VALUE);
                                    oForm.Items.Item("bookpms47").Specific.VALUE = 0;
                                    PSH_Globals.SBO_Application.MessageBox("총급여 7천만원 초과자입니다. 일반금액에 합산하고 도서공연비는 0처리 합니다.");
                                }
                                break;
                            case "bookpms":
                                //도서공연사용분
                                //총급여액계산해서 7천만원이하는 0
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                YEAR_Renamed = oForm.Items.Item("Year").Specific.VALUE;
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                                bookAmt = 0;

                                sQry  = "SELECT SUM(gwase) ";
                                sQry += "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "      Union All ";
                                sQry += "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + YEAR_Renamed + "' + '01' AND '" + YEAR_Renamed + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry += "      Union All";
                                sQry += "      SELECT gwase   = SUM( isnull(a.payrtot1 ,0) + isnull(a.payrtot2,0) + isnull(a.bnstot1,0) + isnull(a.bnstot2,0) )";
                                sQry += "        FROM p_sbservcomp a";
                                sQry += "       WHERE a.saup = '" + CLTCOD + "' ";
                                sQry += "         And a.yyyy   =  '" + YEAR_Renamed + "'";
                                sQry += "         And a.sabun  = '" + MSTCOD + "' ";
                                sQry += "     ) g";

                                oRecordSet.DoQuery(sQry);
                                bookAmt = oRecordSet.Fields.Item(0).Value;
                                //총급여액(과세대상)
                                //7천기준
                                if (bookAmt > 70000000)
                                {
                                    oForm.Items.Item("ntsamt24").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt24").Specific.VALUE) + Conversion.Val(oForm.Items.Item("bookpms").Specific.VALUE);
                                    oForm.Items.Item("ntsamt").Specific.VALUE = Conversion.Val(oForm.Items.Item("ntsamt24").Specific.VALUE);
                                    oForm.Items.Item("bookpms").Specific.VALUE = 0;
                                    PSH_Globals.SBO_Application.MessageBox("총급여 7천만원 초과자입니다. 일반금액에 합산하고 도서공연비는 0처리 합니다.");
                                }
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
                oForm.Freeze(false);
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
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            string Param04 = string.Empty;
            string Param05 = string.Empty;
            string Param06 = string.Empty;
            string Param07 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);

                            Param01 = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                            Param02 = oDS_PH_PY402A.Columns.Item("Year").Cells.Item(pVal.Row).Value.Trim();
                            Param03 = oDS_PH_PY402A.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value.Trim();
                            Param04 = oDS_PH_PY402A.Columns.Item("div").Cells.Item(pVal.Row).Value.Trim();
                            Param05 = oDS_PH_PY402A.Columns.Item("target").Cells.Item(pVal.Row).Value.Trim();
                            Param06 = oDS_PH_PY402A.Columns.Item("relate").Cells.Item(pVal.Row).Value.Trim();
                            Param07 = oDS_PH_PY402A.Columns.Item("juminno").Cells.Item(pVal.Row).Value.Trim();

                            sQry = "EXEC PH_PY402_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "'";
                            oRecordSet.DoQuery(sQry);

                            if ((oRecordSet.RecordCount == 0))
                            {

                                oForm.Items.Item("MSTCOD").Specific.VALUE = oDS_PH_PY402A.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Value;
                                oForm.Items.Item("FullName").Specific.VALUE = oDS_PH_PY402A.Columns.Item("FullName").Cells.Item(pVal.Row).Value;

                                oForm.DataSources.UserDataSources.Item("div").Value = "";
                                oForm.DataSources.UserDataSources.Item("divnm").Value = "";
                                oForm.DataSources.UserDataSources.Item("target").Value = "";
                                oForm.DataSources.UserDataSources.Item("targetnm").Value = "";

                                oForm.Items.Item("relate").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                                oForm.Items.Item("hdcode").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.DataSources.UserDataSources.Item("birthymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("addr").Value = "";

                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("handoamt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ntsamt24").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("mart24").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("trans24").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("bookpms").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ntsamt3").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("mart3").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("trans3").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("bookpms3").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("ntsamt47").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("mart47").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("trans47").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("bookpms47").Value = Convert.ToString(0);

                                oForm.Items.Item("TeamName").Specific.VALUE = "";
                                oForm.Items.Item("RspName").Specific.VALUE = "";
                                oForm.Items.Item("ClsName").Specific.VALUE = "";

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {
                                oForm.Items.Item("Year").Specific.VALUE = oRecordSet.Fields.Item("Year").Value.ToString().Trim();
                                oForm.Items.Item("MSTCOD").Specific.VALUE = oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim();
                                oForm.Items.Item("FullName").Specific.VALUE = oRecordSet.Fields.Item("FullName").Value.ToString().Trim();

                                // 부서
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value.ToString().Trim();

                                oForm.DataSources.UserDataSources.Item("div").Value = oRecordSet.Fields.Item("div").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("divnm").Value = oRecordSet.Fields.Item("divnm").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("target").Value = oRecordSet.Fields.Item("target").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("targetnm").Value = oRecordSet.Fields.Item("targetnm").Value.ToString().Trim();

                                oForm.Items.Item("relate").Specific.Select(oRecordSet.Fields.Item("relate").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                                oForm.Items.Item("hdcode").Specific.Select(oRecordSet.Fields.Item("hdcode").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                                oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("birthymd").Value = oRecordSet.Fields.Item("birthymd").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("addr").Value = oRecordSet.Fields.Item("addr").Value.ToString().Trim();

                                oForm.DataSources.UserDataSources.Item("ntsamt").Value = oRecordSet.Fields.Item("ntsamt").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("amt").Value = oRecordSet.Fields.Item("amt").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("handoamt").Value = oRecordSet.Fields.Item("handoamt").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("ntsamt24").Value = oRecordSet.Fields.Item("ntsamt24").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("mart24").Value = oRecordSet.Fields.Item("mart24").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("trans24").Value = oRecordSet.Fields.Item("trans24").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("bookpms").Value = oRecordSet.Fields.Item("bookpms").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("ntsamt3").Value = oRecordSet.Fields.Item("ntsamt3").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("mart3").Value = oRecordSet.Fields.Item("mart3").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("trans3").Value = oRecordSet.Fields.Item("trans3").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("bookpms3").Value = oRecordSet.Fields.Item("bookpms3").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("ntsamt47").Value = oRecordSet.Fields.Item("ntsamt47").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("mart47").Value = oRecordSet.Fields.Item("mart47").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("trans47").Value = oRecordSet.Fields.Item("trans47").Value.ToString().Trim();
                                oForm.DataSources.UserDataSources.Item("bookpms47").Value = oRecordSet.Fields.Item("bookpms47").Value.ToString().Trim();

                                // 2020
                                if (oRecordSet.Fields.Item("div").Value.ToString().Trim() == "50")
                                {
                                    oForm.Items.Item("ntsamt").Enabled = false;
                                    oForm.Items.Item("amt").Enabled = false;

                                    oForm.Items.Item("ntsamt3").Enabled = true;
                                    oForm.Items.Item("mart3").Enabled = true;
                                    oForm.Items.Item("trans3").Enabled = true;
                                    oForm.Items.Item("bookpms3").Enabled = true;

                                    oForm.Items.Item("ntsamt47").Enabled = true;
                                    oForm.Items.Item("mart47").Enabled = true;
                                    oForm.Items.Item("trans47").Enabled = true;
                                    oForm.Items.Item("bookpms47").Enabled = true;

                                    oForm.Items.Item("ntsamt24").Enabled = true;
                                    oForm.Items.Item("mart24").Enabled = true;
                                    oForm.Items.Item("trans24").Enabled = true;
                                    oForm.Items.Item("bookpms").Enabled = true;
                                }
                                else
                                {
                                    oForm.Items.Item("ntsamt").Enabled = true;
                                    oForm.Items.Item("amt").Enabled = true;

                                    oForm.Items.Item("ntsamt3").Enabled = false;
                                    oForm.Items.Item("mart3").Enabled = false;
                                    oForm.Items.Item("trans3").Enabled = false;
                                    oForm.Items.Item("bookpms3").Enabled = false;

                                    oForm.Items.Item("ntsamt47").Enabled = false;
                                    oForm.Items.Item("mart47").Enabled = false;
                                    oForm.Items.Item("trans47").Enabled = false;
                                    oForm.Items.Item("bookpms47").Enabled = false;

                                    oForm.Items.Item("ntsamt24").Enabled = false;
                                    oForm.Items.Item("mart24").Enabled = false;
                                    oForm.Items.Item("trans24").Enabled = false;
                                    oForm.Items.Item("bookpms").Enabled = false;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY402_DataFind
        /// </summary>
        private void PH_PY402_DataFind()
        {
            int iRow = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string Year = string.Empty;
            string MSTCOD = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
            Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;


            if (string.IsNullOrEmpty(Strings.Trim(CLTCOD)))
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("사업장이 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                return; 
            }

            if (string.IsNullOrEmpty(Strings.Trim(Year)))
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("년도가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                return;
            }

            if (string.IsNullOrEmpty(Strings.Trim(MSTCOD)))
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("사번이 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                return;
            }

            try
            {
                oForm.Freeze(true);

                PH_PY402_FormItemEnabled();

                sQry = "EXEC PH_PY402_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY402A.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY402_TitleSetting(ref iRow);
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY402_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }
        
        /// <summary>
        /// PH_PY402_SAVE
        /// </summary>
        private void PH_PY402_SAVE()
        {
            // 데이타 저장
            short ErrNum = 0;
            string vReturnValue = string.Empty;
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
            string FullName = string.Empty;
            string YEAR = string.Empty;
            string hdcode = string.Empty;
            string Div = string.Empty;
            string target = string.Empty;
            string relate = string.Empty;
            string kname = string.Empty;
            string juminno = string.Empty;
            string addr = string.Empty;
            string birthymd = string.Empty;
            string CheckDate1 = string.Empty;
            string CheckDate2 = string.Empty;

            double Amt = 0;
            double ntsamt = 0;
            double ntsamt24 = 0;
            double mart24 = 0;
            double trans24 = 0;
            double bookpms = 0;
            double ntsamt3 = 0;
            double mart3 = 0;
            double trans3 = 0;
            double bookpms3 = 0;
            double ntsamt47 = 0;
            double mart47 = 0;
            double trans47 = 0;
            double bookpms47 = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                YEAR = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.VALUE.ToString().Trim();

                Div = oForm.Items.Item("div").Specific.VALUE.ToString().Trim();
                target = oForm.Items.Item("target").Specific.VALUE.ToString().Trim();
                relate = oForm.Items.Item("relate").Specific.VALUE.ToString().Trim();
                kname = oForm.Items.Item("kname").Specific.VALUE.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.VALUE.ToString().Trim();
                addr = oForm.Items.Item("addr").Specific.VALUE.ToString().Trim();
                birthymd = oForm.Items.Item("birthymd").Specific.VALUE.ToString().Trim();
                hdcode = oForm.Items.Item("hdcode").Specific.VALUE.ToString().Trim();

                Amt = Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE);
                ntsamt = Convert.ToDouble(oForm.Items.Item("ntsamt").Specific.VALUE);

                ntsamt24 = Convert.ToDouble(oForm.Items.Item("ntsamt24").Specific.VALUE);
                mart24 = Convert.ToDouble(oForm.Items.Item("mart24").Specific.VALUE);
                trans24 = Convert.ToDouble(oForm.Items.Item("trans24").Specific.VALUE);
                bookpms = Convert.ToDouble(oForm.Items.Item("bookpms").Specific.VALUE);

                ntsamt3 = Convert.ToDouble(oForm.Items.Item("ntsamt3").Specific.VALUE);
                mart3 = Convert.ToDouble(oForm.Items.Item("mart3").Specific.VALUE);
                trans3 = Convert.ToDouble(oForm.Items.Item("trans3").Specific.VALUE);
                bookpms3 = Convert.ToDouble(oForm.Items.Item("bookpms3").Specific.VALUE);

                ntsamt47 = Convert.ToDouble(oForm.Items.Item("ntsamt47").Specific.VALUE);
                mart47 = Convert.ToDouble(oForm.Items.Item("mart47").Specific.VALUE);
                trans47 = Convert.ToDouble(oForm.Items.Item("trans47").Specific.VALUE);
                bookpms47 = Convert.ToDouble(oForm.Items.Item("bookpms47").Specific.VALUE);

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
                if (target == "220" && string.IsNullOrWhiteSpace(hdcode))
                {
                    ErrNum = 4;
                    throw new Exception();
                }

                if (Strings.Trim(target) != "220" && !string.IsNullOrEmpty(Strings.Trim(hdcode)))
                {
                    hdcode = "";
                }

                if (string.IsNullOrEmpty(Strings.Trim(juminno)) || (Div != "70"  && Conversion.Val(Amt) + Conversion.Val(ntsamt) + Conversion.Val(ntsamt3) + Conversion.Val(mart3) + Conversion.Val(trans3) + Conversion.Val(bookpms3)
                                                                                                                                 + Conversion.Val(ntsamt47) + Conversion.Val(mart47) + Conversion.Val(trans47) + Conversion.Val(bookpms47)
                                                                                                                                 + Conversion.Val(ntsamt24) + Conversion.Val(mart24) + Conversion.Val(trans24) + Conversion.Val(bookpms) == 0))
                {                                             //기본공제제외자(70)  
                    ErrNum = 5;
                    throw new Exception();
                }

                sQry = " Select U_Char2, U_Char3 From [@PS_HR200L] Where Code = '71' And U_Code = '" + target + "' ";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    CheckDate1 = oRecordSet.Fields.Item(0).Value;
                    CheckDate2 = oRecordSet.Fields.Item(1).Value;

                    if (!string.IsNullOrEmpty(Strings.Trim(CheckDate1)))
                    {
                        if (relate == "05" | relate == "06" | relate == "07" | relate == "12" | relate == "13" | relate == "21" | relate == "22")
                        {
                            if (Convert.ToDouble(birthymd) > Convert.ToDouble(CheckDate1))
                            {
                                vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("기준일자 이후출생자입니다. 계속하겠습니까?", 1, "&확인", "&취소"));
                                switch (vReturnValue)
                                {
                                    case "1":
                                        break;
                                    case "2":
                                        ErrNum = 0;
                                        throw new Exception();
                                }
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(Strings.Trim(CheckDate2)))
                    {
                        if (relate == "03" | relate == "04" | relate == "08" | relate == "23")
                        {
                            if (Convert.ToDouble(birthymd) <= Convert.ToDouble(CheckDate2))
                            {
                                vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("기준일자 이전출생자입니다. 계속하겠습니까?", 1, "&확인", "&취소"));
                                switch (vReturnValue)
                                {
                                    case "1":
                                        break;
                                    case "2":
                                        ErrNum = 0;
                                        throw new Exception();
                                }
                            }
                        }
                    }

                }

                sQry = " Select Count(*) From [p_seoybase] Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                sQry = sQry + " And div = '" + Div + "' And target = '" + target + "' And relate = '" + relate + "' And juminno = '" + juminno + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신
                    sQry  = "Update [p_seoybase] set ";
                    sQry += "kname = '" + kname + "',";
                    sQry += "addr = '" + addr + "',";
                    sQry += "birthymd = '" + birthymd + "',";
                    sQry += "hdcode = '" + hdcode + "',";
                    sQry += "amt = " + Amt + ",";
                    sQry += "ntsamt =" + ntsamt + ",";
                    sQry += "ntsamt3 =" + ntsamt3 + ",";
                    sQry += "mart3 =" + mart3 + ",";
                    sQry += "trans3 =" + trans3 + ",";
                    sQry += "bookpms3 =" + bookpms3 + ",";
                    sQry += "ntsamt47 =" + ntsamt47 + ",";
                    sQry += "mart47 =" + mart47 + ",";
                    sQry += "trans47 =" + trans47 + ",";
                    sQry += "bookpms47 =" + bookpms47 + ",";
                    sQry += "ntsamt24 =" + ntsamt24 + ",";
                    sQry += "mart24 =" + mart24 + ",";
                    sQry += "trans24 =" + trans24 + ",";
                    sQry += "bookpms =" + bookpms;
                    sQry += " Where saup = '" + CLTCOD + "' And yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                    sQry += " And div = '" + Div + "' And target = '" + target + "' And relate = '" + relate + "' And juminno = '" + juminno + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY402_DataFind();
                }
                else
                {
                    // 신규
                    sQry  = "INSERT INTO [p_seoybase]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "div,";
                    sQry += "target,";
                    sQry += "relate,";
                    sQry += "kname,";
                    sQry += "juminno,";
                    sQry += "addr,";
                    sQry += "birthymd,";
                    sQry += "hdcode,";
                    sQry += "amt,";
                    sQry += "ntsamt,";
                    sQry += "soduk,";
                    sQry += "ntsamt24,";
                    sQry += "mart24, ";
                    sQry += "trans24, ";
                    sQry += "bookpms, ";
                    sQry += "ntsamt3,";
                    sQry += "mart3, ";
                    sQry += "trans3, ";
                    sQry += "bookpms3, ";
                    sQry += "ntsamt47,";
                    sQry += "mart47, ";
                    sQry += "trans47, ";
                    sQry += "bookpms47 )";
                    sQry += " VALUES(";

                    sQry += "'" + CLTCOD + "',";
                    sQry += "'" + YEAR + "',";
                    sQry += "'" + MSTCOD + "',";
                    sQry += "'" + Div + "',";
                    sQry += "'" + target + "',";
                    sQry += "'" + relate + "',";
                    sQry += "'" + kname + "',";
                    sQry += "'" + juminno + "',";
                    sQry += "'" + addr + "',";
                    sQry += "'" + birthymd + "',";
                    sQry += "'" + hdcode + "',";
                    sQry += Amt + ",";
                    sQry += ntsamt + ", 0 ,";
                    sQry += ntsamt24 + ",";
                    sQry += mart24 + ",";
                    sQry += trans24 + ",";
                    sQry += bookpms + ",";
                    sQry += ntsamt3 + ",";
                    sQry += mart3 + ",";
                    sQry += trans3 + ",";
                    sQry += bookpms3 + ",";
                    sQry += ntsamt47 + ",";
                    sQry += mart47 + ",";
                    sQry += trans47 + ",";
                    sQry += bookpms47 + " )";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY402_DataFind();
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
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.MessageBox("장애인코드가 없습니다. 장애인인 경우 장애인 코드를 선택바랍니다. 확인바랍니다.");
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY402_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY402_Delete  데이타 삭제
        /// </summary>
        private void PH_PY402_Delete()
        {
            string CLTCOD = string.Empty;
            string YEAR = string.Empty;
            string MSTCOD = string.Empty;
            string Div = string.Empty;
            string target = string.Empty;
            string relate = string.Empty;
            string juminno = string.Empty;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                YEAR = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                Div = oForm.Items.Item("div").Specific.VALUE.ToString().Trim();
                target = oForm.Items.Item("target").Specific.VALUE.ToString().Trim();
                relate = oForm.Items.Item("relate").Specific.VALUE.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.VALUE.ToString().Trim();

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY402A.Rows.Count > 0)
                    {
                        sQry  = "Delete From [p_seoybase] Where saup = '" + CLTCOD + "' AND  yyyy = '" + YEAR + "' And sabun = '" + MSTCOD + "'";
                        sQry += " And div = '" + Div + "' And target = '" + target + "' And relate = '" + relate + "' And juminno = '" + juminno + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY402_DataFind();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY402_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY402_TitleSetting
        /// </summary>
        private void PH_PY402_TitleSetting(ref int iRow)
        {
            int i = 0;
            string[] COLNAM = new string[16];

            try
            {

                COLNAM[0] = "년도";
                COLNAM[1] = "사번";
                COLNAM[2] = "공제구분코드";
                COLNAM[3] = "공제구분";
                COLNAM[4] = "공제대상코드";
                COLNAM[5] = "공제대상";
                COLNAM[6] = "관계코드";
                COLNAM[7] = "관계";
                COLNAM[8] = "성명";
                COLNAM[9] = "주민번호";
                COLNAM[10] = "금액(국세청)";
                COLNAM[11] = "금액(국세청외)";
                COLNAM[12] = "전통시장";
                COLNAM[13] = "대중교통";
                COLNAM[14] = "도서공연";
                COLNAM[15] = "합계금액";

                for (i = 0; i <= Information.UBound(COLNAM); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    oGrid1.Columns.Item(i).Editable = false;
                    if (COLNAM[i] == "사번" | COLNAM[i] == "공제구분코드" | COLNAM[i] == "공제대상코드" | COLNAM[i] == "관계코드" | COLNAM[i] == "주민번호")
                    {
                        oGrid1.Columns.Item(i).Visible = false;
                    }
                    oGrid1.Columns.Item(i).RightJustified = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY402_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }
    }
}
