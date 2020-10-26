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
    /// 의료비자료등록
    /// </summary>
    internal class PH_PY405 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY405;
        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY405L;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY405.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY405_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY405");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY405_CreateItems();
                PH_PY405_FormItemEnabled();
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
        private void PH_PY405_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY405");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY405");
                oDS_PH_PY405 = oForm.DataSources.DataTables.Item("PH_PY405");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oDS_PH_PY405L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("관계코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("관계명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("주민번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급처상호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("사업자번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급일자", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("의료증빙코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("의료증빙명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급금액(외)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급금액(국)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("건수", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("경로", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("장애", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("난임", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("특례", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

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

                // 관계
                oForm.DataSources.UserDataSources.Add("rel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("rel").Specific.DataBind.SetBound(true, "", "rel");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("rel").Specific, "Y");
                oForm.Items.Item("rel").DisplayDesc = true;
                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 성명
                oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

                // 주민등록번호
                oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

                // 내.외국인코드
                oForm.DataSources.UserDataSources.Add("empdiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("empdiv").Specific.DataBind.SetBound(true, "", "empdiv");
                oForm.Items.Item("empdiv").Specific.ValidValues.Add("1", "내국인");
                oForm.Items.Item("empdiv").Specific.ValidValues.Add("9", "외국인");
                oForm.Items.Item("empdiv").DisplayDesc = true;
                oForm.Items.Item("empdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 지급처상호
                oForm.DataSources.UserDataSources.Add("custnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("custnm").Specific.DataBind.SetBound(true, "", "custnm");

                // 사업자등록번호
                oForm.DataSources.UserDataSources.Add("entno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("entno").Specific.DataBind.SetBound(true, "", "entno");

                // 지급일자
                oForm.DataSources.UserDataSources.Add("payymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("payymd").Specific.DataBind.SetBound(true, "", "payymd");

                // 의료증빙코드
                oForm.DataSources.UserDataSources.Add("gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("gubun").Specific.DataBind.SetBound(true, "", "gubun");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("1", "국세청장이제공하는의료비자료");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("2", "국민건강보험공단의의료비부담명세서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("3", "진료비계산서,약제비계산서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("4", "장기요양급여비용명세서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("5", "기타의료비영수증");
                oForm.Items.Item("gubun").DisplayDesc = true;
                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 지급금액(국세청자료외)
                oForm.DataSources.UserDataSources.Add("medcex", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcex").Specific.DataBind.SetBound(true, "", "medcex");

                // 지급금액(국세청자료)
                oForm.DataSources.UserDataSources.Add("ntamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntamt").Specific.DataBind.SetBound(true, "", "ntamt");

                // 지급건수
                oForm.DataSources.UserDataSources.Add("cont", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("cont").Specific.DataBind.SetBound(true, "", "cont");

                // 경로여부
                oForm.DataSources.UserDataSources.Add("olddiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("olddiv").Specific.DataBind.SetBound(true, "", "olddiv");
                oForm.Items.Item("olddiv").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("olddiv").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("olddiv").DisplayDesc = true;
                oForm.Items.Item("olddiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 장애여부
                oForm.DataSources.UserDataSources.Add("deform", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("deform").Specific.DataBind.SetBound(true, "", "deform");
                oForm.Items.Item("deform").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("deform").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("deform").DisplayDesc = true;
                oForm.Items.Item("deform").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 난임시술비여부
                oForm.DataSources.UserDataSources.Add("nanim", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("nanim").Specific.DataBind.SetBound(true, "", "nanim");
                oForm.Items.Item("nanim").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("nanim").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("nanim").DisplayDesc = true;
                oForm.Items.Item("nanim").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 건겅보험산정특례자여부
                oForm.DataSources.UserDataSources.Add("tukrae", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("tukrae").Specific.DataBind.SetBound(true, "", "tukrae");
                oForm.Items.Item("tukrae").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("tukrae").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("tukrae").DisplayDesc = true;
                oForm.Items.Item("tukrae").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY405_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oForm.EnableMenu("1282", true);  // 문서추가

                if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE)))
                {
                    oForm.Items.Item("Year").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);
                }
                //oForm.Items.Item("MSTCOD").Specific.VALUE = "";
                //oForm.Items.Item("FullName").Specific.VALUE = "";
                //oForm.Items.Item("TeamName").Specific.VALUE = "";
                //oForm.Items.Item("RspName").Specific.VALUE = "";
                //oForm.Items.Item("ClsName").Specific.VALUE = "";

                //oForm.Items("kname").Specific.VALUE = ""
                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                oForm.Items.Item("juminno").Specific.VALUE = "";
                oForm.Items.Item("custnm").Specific.VALUE = "";
                oForm.Items.Item("entno").Specific.VALUE = "";
                oForm.Items.Item("payymd").Specific.VALUE = "";

                oForm.Items.Item("medcex").Specific.VALUE = 0;
                oForm.Items.Item("ntamt").Specific.VALUE = 0;
                oForm.Items.Item("cont").Specific.VALUE = 0;

                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("olddiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("deform").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("nanim").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("tukrae").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                ////Key set
                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;

                oForm.Items.Item("juminno").Enabled = true;
                oForm.Items.Item("custnm").Enabled = true;
                oForm.Items.Item("payymd").Enabled = true;
                oForm.Items.Item("entno").Enabled = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY405_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY405);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY405L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                            PH_PY405_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY405_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY405_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY405_FormItemEnabled();
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
                        PH_PY405_DataFind();
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
                            PH_PY405_SAVE();
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
                            PH_PY405_Delete();
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
            string CLTCOD = string.Empty;
            string MSTCOD = string.Empty;
            string yyyy = string.Empty;
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
                        //if (pVal.ItemUID == "rel")
                        //{
                        //    CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                        //    MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                        //    yyyy = oForm.Items.Item("Year").Specific.VALUE;

                        //    if (!string.IsNullOrEmpty(oForm.Items.Item("rel").Specific.VALUE))
                        //    {
                        //        oForm.DataSources.UserDataSources.Item("kname").Value = "";
                        //        oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                        //    }

                        //    sQry = "Select Distinct kname, juminno ";
                        //    sQry = sQry + " From [p_seoybase]";
                        //    sQry = sQry + " Where saup = '" + CLTCOD + "'";
                        //    sQry = sQry + " and sabun = '" + MSTCOD + "'";
                        //    sQry = sQry + " and div In ('10','70') ";
                        //    sQry = sQry + " and relate = '" + oForm.Items.Item("rel").Specific.VALUE + "'";
                        //    sQry = sQry + " and yyyy = '" + yyyy + "'";

                        //    oRecordSet.DoQuery(sQry);

                        //    if (oRecordSet.RecordCount == 1)
                        //    {
                        //        oForm.Items.Item("kname").Specific.VALUE = oRecordSet.Fields.Item("kname").Value;
                        //        oForm.Items.Item("juminno").Specific.VALUE = oRecordSet.Fields.Item("juminno").Value;
                        //    }
                        //}
                        if (pVal.ItemUID == "rel")
                        {
                            oMat01.Clear();
                            oDS_PH_PY405L.Clear();

                            //MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                            //relate = oForm.Items.Item("rel").Specific.VALUE;

                            //sQry = "EXEC [PH_PY407_03] '" + MSTCOD + "', '" + relate + "'";

                            //oRecordSet.DoQuery(sQry);
                            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                            yyyy = oForm.Items.Item("Year").Specific.VALUE;

                            if (!string.IsNullOrEmpty(oForm.Items.Item("rel").Specific.VALUE))
                            {
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                            }

                            sQry = "Select Distinct kname, juminno, birthymd, relatenm = ( select U_CodeNm From[@PS_HR200L] WHERE Code = 'P121' AND U_Code = relate) ";
                            sQry = sQry + " From [p_seoybase]";
                            sQry = sQry + " Where saup = '" + CLTCOD + "'";
                            sQry = sQry + " and sabun = '" + MSTCOD + "'";
                            sQry = sQry + " and div In ('10','70') ";
                            sQry = sQry + " and relate = '" + oForm.Items.Item("rel").Specific.VALUE + "'";
                            sQry = sQry + " and yyyy = '" + yyyy + "'";

                            oRecordSet.DoQuery(sQry);

                            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                            {
                                if (i + 1 > oDS_PH_PY405L.Size)
                                {
                                    oDS_PH_PY405L.InsertRecord((i));
                                }

                                oMat01.AddRow();
                                oDS_PH_PY405L.Offset = i;

                                oDS_PH_PY405L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                oDS_PH_PY405L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet.Fields.Item("kname").Value));
                                oDS_PH_PY405L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet.Fields.Item("juminno").Value));
                                oDS_PH_PY405L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet.Fields.Item("birthymd").Value));
                                oDS_PH_PY405L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet.Fields.Item("relatenm").Value));
                                oRecordSet.MoveNext();
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                  
                            if ((oRecordSet.RecordCount == 1))
                            {
                                oForm.Items.Item("kname").Specific.VALUE = oMat01.Columns.Item("kname").Cells.Item(1).Specific.VALUE;
                                oForm.Items.Item("juminno").Specific.VALUE = oMat01.Columns.Item("juminno").Cells.Item(1).Specific.VALUE;
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
            string CLTCOD, MSTCOD, FullName, rel, kname, yyyy, juminno = string.Empty;
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
                            case "FullName":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                FullName = oForm.Items.Item("FullName").Specific.VALUE;

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
                                sQry = sQry + " And U_status <> '5'"; // 퇴사자 제외
                                sQry = sQry + " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "kname":
                                CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                                rel = oForm.Items.Item("rel").Specific.VALUE;
                                kname = oForm.Items.Item("kname").Specific.VALUE;
                                yyyy = oForm.Items.Item("Year").Specific.VALUE;

                                oForm.Items.Item("juminno").Specific.VALUE = "";

                                sQry = "Select Distinct juminno ";
                                sQry = sQry + " From [p_seoybase]";
                                sQry = sQry + " Where saup = '" + CLTCOD + "'";
                                sQry = sQry + " and sabun = '" + MSTCOD + "'";
                                sQry = sQry + " and relate = '" + rel + "'";
                                sQry = sQry + " and kname = '" + kname + "'";
                                sQry = sQry + " and yyyy = '" + yyyy + "'";

                                oRecordSet.DoQuery(sQry);

                                juminno = oRecordSet.Fields.Item("juminno").Value;
                                if (!string.IsNullOrEmpty(Strings.Trim(juminno)))
                                {
                                    oForm.Items.Item("juminno").Specific.VALUE = juminno;


                                    if (rel != "01")
                                    {
                                        // 65세 경로우대 의료비 체크
                                        sQry = "select Cnt = Count(*) from p_seoybase a ";
                                        sQry = sQry + " Where a.yyyy = '" + yyyy + "'";
                                        sQry = sQry + " and datediff(yy, Left(a.birthymd,4) + '1231', '" + yyyy + "1231'" + " ) >= 65";
                                        sQry = sQry + " And a.juminno = '" + juminno + "'";
                                        oRecordSet.DoQuery(sQry);

                                        if (oRecordSet.Fields.Item("Cnt").Value > 0)
                                        {
                                            oForm.Items.Item("olddiv").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        else
                                        {
                                            oForm.Items.Item("olddiv").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }

                                        // 장애자인경우
                                        sQry = " Select Cnt = Count(*) From p_seoybase ";
                                        sQry = sQry + " Where yyyy = '" + yyyy + "'";
                                        sQry = sQry + " and div = '20' and target = '220'";
                                        sQry = sQry + " And juminno = '" + juminno + "'";
                                        oRecordSet.DoQuery(sQry);

                                        if (oRecordSet.Fields.Item("Cnt").Value > 0)
                                        {
                                            oForm.Items.Item("deform").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        else
                                        {
                                            oForm.Items.Item("deform").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                    }
                                    else
                                    {
                                        oForm.Items.Item("olddiv").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("deform").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("nanim").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("tukrae").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.SetStatusBarMessage("기본공제대상자가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    return;
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
                string Param01, Param02, Param03, Param04, Param05, Param06, Param07 = string.Empty;

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);
                            Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                            Param02 = oDS_PH_PY405.Columns.Item("연도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY405.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY405.Columns.Item("주민번호").Cells.Item(pVal.Row).Value;
                            Param05 = oDS_PH_PY405.Columns.Item("지급처상호").Cells.Item(pVal.Row).Value;
                            Param06 = oDS_PH_PY405.Columns.Item("지급일자").Cells.Item(pVal.Row).Value;
                            Param07 = oDS_PH_PY405.Columns.Item("사업자번호").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY405_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "'";
                            oRecordSet.DoQuery(sQry);

                            if ((oRecordSet.RecordCount == 0))
                            {

                                //oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY405A.Columns.Item("MSTCOD").Cells(oRow).VALUE
                                //oForm.Items("FullName").Specific.VALUE = oDS_PH_PY405A.Columns.Item("FullName").Cells(oRow).VALUE

                                oForm.Items.Item("kname").Specific.VALUE = "";
                                oForm.Items.Item("juminno").Specific.VALUE = "";
                                oForm.Items.Item("custnm").Specific.VALUE = "";
                                oForm.Items.Item("entno").Specific.VALUE = "";
                                oForm.Items.Item("payymd").Specific.VALUE = "";

                                oForm.Items.Item("medcex").Specific.VALUE = 0;
                                oForm.Items.Item("ntamt").Specific.VALUE = 0;
                                oForm.Items.Item("cont").Specific.VALUE = 0;

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {

                                oForm.Items.Item("rel").Specific.Select(oRecordSet.Fields.Item("rel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value;
                                oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value;

                                oForm.Items.Item("empdiv").Specific.Select(oRecordSet.Fields.Item("empdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("custnm").Value = oRecordSet.Fields.Item("custnm").Value;
                                oForm.DataSources.UserDataSources.Item("entno").Value = oRecordSet.Fields.Item("entno").Value;
                                oForm.DataSources.UserDataSources.Item("payymd").Value = oRecordSet.Fields.Item("payymd").Value;

                                oForm.Items.Item("gubun").Specific.Select(oRecordSet.Fields.Item("gubun").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("medcex").Value = oRecordSet.Fields.Item("medcex").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ntamt").Value = oRecordSet.Fields.Item("ntamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("cont").Value = oRecordSet.Fields.Item("cont").Value.ToString();

                                oForm.Items.Item("olddiv").Specific.Select(oRecordSet.Fields.Item("olddiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("deform").Specific.Select(oRecordSet.Fields.Item("deform").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("nanim").Specific.Select(oRecordSet.Fields.Item("nanim").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("tukrae").Specific.Select(oRecordSet.Fields.Item("tukrae").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                //    '//부서
                                //    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                                //    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                                //    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE

                                ////Key Disable
                                oForm.Items.Item("CLTCOD").Enabled = false;
                                oForm.Items.Item("Year").Enabled = false;
                                oForm.Items.Item("MSTCOD").Enabled = false;

                                oForm.Items.Item("juminno").Enabled = false;
                                oForm.Items.Item("custnm").Enabled = false;
                                oForm.Items.Item("payymd").Enabled = false;
                                oForm.Items.Item("entno").Enabled = false;
                            }
                        }
                    }

                    if (pVal.ItemUID == "Mat01")
                    {
                        oForm.Items.Item("kname").Specific.VALUE = oMat01.Columns.Item("kname").Cells.Item(pVal.Row).Specific.VALUE;
                        oForm.Items.Item("juminno").Specific.VALUE = oMat01.Columns.Item("juminno").Cells.Item(pVal.Row).Specific.VALUE;
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
        /// PH_PY405_DataFind
        /// </summary>
        private void PH_PY405_DataFind()
        {
            string sQry = string.Empty;
            string CLTCOD, Year, MSTCOD = string.Empty;
            
            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
            Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

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

                PH_PY405_FormItemEnabled();

                sQry = "EXEC PH_PY405_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY405.ExecuteQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY405_SAVE
        /// </summary>
        private void PH_PY405_SAVE()
        {
            // 데이타 저장
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string saup, yyyy, sabun, kname, juminno, custnm, payymd, rel, empdiv, entno, Gubun, olddiv, deform, nanim, tukrae = string.Empty;
            double medcex, ntamt, cont = 0;

            try
            {
                oForm.Freeze(true);

                saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                yyyy = oForm.Items.Item("Year").Specific.VALUE;
                sabun = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);

                rel = oForm.Items.Item("rel").Specific.VALUE;
                kname = oForm.Items.Item("kname").Specific.VALUE;
                juminno = oForm.Items.Item("juminno").Specific.VALUE;
                empdiv = oForm.Items.Item("empdiv").Specific.VALUE;
                custnm = oForm.Items.Item("custnm").Specific.VALUE;
                entno = oForm.Items.Item("entno").Specific.VALUE;
                payymd = oForm.Items.Item("payymd").Specific.VALUE;
                Gubun = oForm.Items.Item("gubun").Specific.VALUE;
                medcex = Convert.ToDouble(oForm.Items.Item("medcex").Specific.VALUE);
                ntamt = Convert.ToDouble(oForm.Items.Item("ntamt").Specific.VALUE);
                cont = Convert.ToDouble(oForm.Items.Item("cont").Specific.VALUE);
                olddiv = oForm.Items.Item("olddiv").Specific.VALUE;
                deform = oForm.Items.Item("deform").Specific.VALUE;
                nanim = oForm.Items.Item("nanim").Specific.VALUE;
                tukrae = oForm.Items.Item("tukrae").Specific.VALUE;

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

                if (Strings.Trim(olddiv) == "Y" & Strings.Trim(deform) == "Y")
                {
                    PSH_Globals.SBO_Application.MessageBox("경로여부와 장애여부는 둘다'Y'일 수 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(Strings.Trim(juminno)) | (medcex == 0 & ntamt == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                if (medcex != 0 & ntamt != 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("국세청자료와 국세청자료외는 구분하여 별도로 입력 하십시요. 확인바랍니다..");
                    return;
                }

                if (medcex != 0)
                {
                    if (string.IsNullOrEmpty(entno))
                    {
                        PSH_Globals.SBO_Application.MessageBox("사업자등록번호를 확인바랍니다..");
                        return;
                    }
                    if (cont == 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("지급건수를 확인바랍니다..");
                        return;
                    }
                }

                sQry = " Select Count(*) From [p_seoymedhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신

                    sQry = "Update [p_seoymedhis] set ";
                    sQry = sQry + "rel = '" + rel + "',";
                    sQry = sQry + "kname = '" + kname + "',";
                    sQry = sQry + "juminno = '" + juminno + "',";
                    sQry = sQry + "empdiv = '" + empdiv + "',";
                    sQry = sQry + "custnm = '" + custnm + "',";
                    sQry = sQry + "entno = '" + entno + "',";
                    sQry = sQry + "payymd = '" + payymd + "',";
                    sQry = sQry + "gubun = '" + Gubun + "',";

                    sQry = sQry + "medcex = " + medcex + ",";
                    sQry = sQry + "ntamt = " + ntamt + ",";
                    sQry = sQry + "cont = " + cont + ",";

                    sQry = sQry + "olddiv = '" + olddiv + "',";
                    sQry = sQry + "deform = '" + deform + "',";
                    sQry = sQry + "tukrae = '" + tukrae + "',";
                    sQry = sQry + "nanim = '" + nanim + "'";

                    sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY405_DataFind();
                }
                else
                {
                    // 신규
                    sQry = "INSERT INTO [p_seoymedhis]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "rel,";
                    sQry = sQry + "kname,";
                    sQry = sQry + "juminno,";
                    sQry = sQry + "empdiv,";
                    sQry = sQry + "custnm,";
                    sQry = sQry + "entno,";
                    sQry = sQry + "payymd,";
                    sQry = sQry + "gubun,";
                    sQry = sQry + "medcex,";
                    sQry = sQry + "ntamt,";
                    sQry = sQry + "cont,";
                    sQry = sQry + "olddiv,";
                    sQry = sQry + "deform,";
                    sQry = sQry + "nanim,";
                    sQry = sQry + "tukrae,";
                    sQry = sQry + "mednm";
                    sQry = sQry + " ) ";
                    sQry = sQry + "VALUES(";

                    sQry = sQry + "'" + saup + "',";
                    sQry = sQry + "'" + yyyy + "',";
                    sQry = sQry + "'" + sabun + "',";
                    sQry = sQry + "'" + rel + "',";
                    sQry = sQry + "'" + kname + "',";
                    sQry = sQry + "'" + juminno + "',";
                    sQry = sQry + "'" + empdiv + "',";
                    sQry = sQry + "'" + custnm + "',";
                    sQry = sQry + "'" + entno + "',";
                    sQry = sQry + "'" + payymd + "',";
                    sQry = sQry + "'" + Gubun + "',";

                    sQry = sQry + medcex + ",";
                    sQry = sQry + ntamt + ",";
                    sQry = sQry + cont + ",";

                    sQry = sQry + "'" + olddiv + "',";
                    sQry = sQry + "'" + deform + "',";
                    sQry = sQry + "'" + nanim + "',";
                    sQry = sQry + "'" + tukrae + "',";
                    sQry = sQry + "'" + "" + "'";
                    sQry = sQry + " ) ";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY405_DataFind();
                }
            }
            catch (Exception ex)
            {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY405_Delete
        /// </summary>
        private void PH_PY405_Delete()
        {
            // 데이타 삭제
            string sQry = string.Empty;
            string saup, yyyy, sabun, juminno, custnm, entno, payymd = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                yyyy = oForm.Items.Item("Year").Specific.VALUE;
                sabun = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);
                juminno = oForm.Items.Item("juminno").Specific.VALUE;
                custnm = oForm.Items.Item("custnm").Specific.VALUE;
                entno = oForm.Items.Item("entno").Specific.VALUE;
                payymd = oForm.Items.Item("payymd").Specific.VALUE;

                
                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY405.Rows.Count > 0)
                    {
                        sQry = "Delete From [p_seoymedhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY405_DataFind();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }
    }
}
