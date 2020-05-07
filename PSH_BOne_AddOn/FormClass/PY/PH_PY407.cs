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
    /// 정산기부금등록
    /// </summary>
    internal class PH_PY407 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        // 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.Matrix oMat01;
        public SAPbouiCOM.DataTable oDS_PH_PY407A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY407L;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY407.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY407_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY407");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY407_CreateItems();
                PH_PY407_FormItemEnabled();
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
        private void PH_PY407_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY407L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oGrid1 = oForm.Items.Item("Grid01").Specific;

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oForm.DataSources.DataTables.Add("PH_PY407");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY407");
                oDS_PH_PY407A = oForm.DataSources.DataTables.Item("PH_PY407");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("연도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("관계", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("관계명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("주민번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부내용", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사업자번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부처명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금액(국세청)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금액(국세청외)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부장려금신청금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                
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

                //성명
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
                oForm.Items.Item("rel").Specific.ValidValues.Add("", "");
                oForm.Items.Item("rel").Specific.ValidValues.Add("1", "거주자");
                oForm.Items.Item("rel").Specific.ValidValues.Add("2", "배우자");
                oForm.Items.Item("rel").Specific.ValidValues.Add("3", "직계비속");
                oForm.Items.Item("rel").Specific.ValidValues.Add("4", "직계존속");
                oForm.Items.Item("rel").Specific.ValidValues.Add("5", "형제,자매");
                oForm.Items.Item("rel").Specific.ValidValues.Add("6", "그외");
                oForm.Items.Item("rel").DisplayDesc = true;
                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                
                // 성명
                oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

                // 주민번호
                oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

                // 기부금코드  73
                oForm.DataSources.UserDataSources.Add("gibucd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("gibucd").Specific, "Y");

                // 기부내용 2018추가
                oForm.DataSources.UserDataSources.Add("gibudscr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("gibudscr").Specific.ValidValues.Add("1", "금전");
                oForm.Items.Item("gibudscr").Specific.ValidValues.Add("2", "현물");
                oForm.Items.Item("gibudscr").DisplayDesc = true;
                oForm.Items.Item("gibudscr").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 사업자번호
                oForm.DataSources.UserDataSources.Add("saupno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("saupno").Specific.DataBind.SetBound(true, "", "saupno");

                // 부처명
                oForm.DataSources.UserDataSources.Add("sangho", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("sangho").Specific.DataBind.SetBound(true, "", "sangho");

                // 공제금액(국세청)
                oForm.DataSources.UserDataSources.Add("ntamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntamt").Specific.DataBind.SetBound(true, "", "ntamt");

                // 공제금액(국세청외)
                oForm.DataSources.UserDataSources.Add("amt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("amt").Specific.DataBind.SetBound(true, "", "amt");

                // 기부장려금신청금액 2016
                oForm.DataSources.UserDataSources.Add("jamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("jamt").Specific.DataBind.SetBound(true, "", "jamt");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY407_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                oForm.EnableMenu("1282", true);      // 문서추가

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

                oForm.Items.Item("rel").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("gibudscr").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                oForm.DataSources.UserDataSources.Item("sangho").Value = "";

                oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);

                // Key set
                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;

                oForm.Items.Item("juminno").Enabled = true;
                oForm.Items.Item("saupno").Enabled = true;
                oForm.Items.Item("gibucd").Enabled = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY407_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY407L);
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
                            PH_PY407_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY407_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY407_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY407_FormItemEnabled();
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
                        PH_PY407_DataFind();
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
                            PH_PY407_SAVE();
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
                            PH_PY407_Delete();
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
            string MSTCOD, relate = string.Empty;
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
                        if (pVal.ItemUID == "rel")
                        {
                            oMat01.Clear();
                            oDS_PH_PY407L.Clear();

                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
                            relate = oForm.Items.Item("rel").Specific.VALUE;

                            sQry = "EXEC [PH_PY407_03] '" + MSTCOD + "', '" + relate + "'";

                            oRecordSet.DoQuery(sQry);

                            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                            {
                                if (i + 1 > oDS_PH_PY407L.Size)
                                {
                                    oDS_PH_PY407L.InsertRecord((i));
                                }

                                oMat01.AddRow();
                                oDS_PH_PY407L.Offset = i;

                                oDS_PH_PY407L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                oDS_PH_PY407L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet.Fields.Item("kname").Value));
                                oDS_PH_PY407L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet.Fields.Item("juminno").Value));
                                oDS_PH_PY407L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet.Fields.Item("birthymd").Value));
                                oDS_PH_PY407L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet.Fields.Item("relate").Value));
                                oDS_PH_PY407L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet.Fields.Item("addr").Value));
                                oRecordSet.MoveNext();
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();

                            if ((oRecordSet.RecordCount == 0))
                            {
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                                oForm.DataSources.UserDataSources.Item("sangho").Value = "";

                                oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);

                            }

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
            string CLTCOD, MSTCOD, FullName, Div, target, YEAR_Renamed = string.Empty;
            Double bookAmt = 0;
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
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;

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
                                sQry = sQry + " And U_status <> '5'";
                                //퇴사자 제외
                                sQry = sQry + " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                oForm.Items.Item("TeamName").Specific.VALUE = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.VALUE = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.VALUE = oRecordSet.Fields.Item("ClsName").Value;
                                break;
                            case "ntamt":
                                break;

                            case "juminno":
                                //주민번호
                                //주민번호입력시 생년월일 생성
                                if (Strings.Len(Strings.Trim(oForm.Items.Item("juminno").Specific.VALUE)) != 13)
                                {
                                    PSH_Globals.SBO_Application.MessageBox("주민번호자릿수가 틀립니다. 확인하세요.");
                                }
                                else
                                {
                                }
                                break;
                            case "amt":
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
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string Param01, Param02, Param03, Param04, Param05, Param06 = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);

                            Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
                            Param02 = oDS_PH_PY407A.Columns.Item("연도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY407A.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY407A.Columns.Item("사업자번호").Cells.Item(pVal.Row).Value;
                            Param05 = oDS_PH_PY407A.Columns.Item("기부금코드").Cells.Item(pVal.Row).Value;
                            Param06 = oDS_PH_PY407A.Columns.Item("주민번호").Cells.Item(pVal.Row).Value;


                            sQry = "EXEC PH_PY407_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "'";
                            oRecordSet.DoQuery(sQry);

                            if ((oRecordSet.RecordCount == 0))
                            {
                                oForm.Items.Item("rel").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("gibudscr").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                                oForm.DataSources.UserDataSources.Item("sangho").Value = "";
                                oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
                                oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }

                            oForm.Items.Item("gibucd").Specific.Select(oRecordSet.Fields.Item("gibucd").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("gibudscr").Specific.Select(oRecordSet.Fields.Item("gibudscr").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("rel").Specific.Select(oRecordSet.Fields.Item("rel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value;
                            oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value;
                            oForm.DataSources.UserDataSources.Item("saupno").Value = oRecordSet.Fields.Item("saupno").Value;
                            oForm.DataSources.UserDataSources.Item("sangho").Value = oRecordSet.Fields.Item("sangho").Value;
                            oForm.DataSources.UserDataSources.Item("ntamt").Value = oRecordSet.Fields.Item("ntamt").Value.ToString();
                            oForm.DataSources.UserDataSources.Item("amt").Value = oRecordSet.Fields.Item("amt").Value.ToString();
                            oForm.DataSources.UserDataSources.Item("jamt").Value = oRecordSet.Fields.Item("jamt").Value.ToString();

                            oForm.ActiveItem = "rel";

                            ////key set
                            oForm.Items.Item("CLTCOD").Enabled = false;
                            oForm.Items.Item("Year").Enabled = false;
                            oForm.Items.Item("MSTCOD").Enabled = false;

                            oForm.Items.Item("gibucd").Enabled = false;
                            oForm.Items.Item("juminno").Enabled = false;
                            oForm.Items.Item("saupno").Enabled = false;
                            
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
        /// PH_PY407_DataFind
        /// </summary>
        private void PH_PY407_DataFind()
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

                PH_PY407_FormItemEnabled();

                sQry = "EXEC PH_PY407_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY407A.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY407_TitleSetting(ref iRow);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_SAVE
        /// </summary>
        private void PH_PY407_SAVE()
        {
            // 데이타 저장
            short ErrNum = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string vReturnValue = string.Empty;
            string saup, yyyy, sabun, rel, kname, juminno = string.Empty;
            string gibucd, gibudscr, saupno, sangho, FullName = string.Empty;

            double Amt, ntamt, jamt = 0;

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                rel = oForm.Items.Item("rel").Specific.VALUE.ToString().Trim();
                kname = oForm.Items.Item("kname").Specific.VALUE.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.VALUE.ToString().Trim();
                gibucd = oForm.Items.Item("gibucd").Specific.VALUE.ToString().Trim();
                gibudscr = oForm.Items.Item("gibudscr").Specific.VALUE.ToString().Trim();
                saupno = oForm.Items.Item("saupno").Specific.VALUE.ToString().Trim();
                sangho = oForm.Items.Item("sangho").Specific.VALUE.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.VALUE.ToString().Trim();
                ntamt = Convert.ToDouble(oForm.Items.Item("ntamt").Specific.VALUE);
                Amt = Convert.ToDouble(oForm.Items.Item("amt").Specific.VALUE);
                jamt = Convert.ToDouble(oForm.Items.Item("jamt").Specific.VALUE);
                
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
                if (string.IsNullOrEmpty(Strings.Trim(juminno)) | (ntamt == 0 & Amt == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신
                    sQry = "Update [p_seoygibuhis] set ";
                    sQry = sQry + "rel = '" + rel + "',";
                    sQry = sQry + "kname = '" + kname + "',";
                    sQry = sQry + "sangho = '" + sangho + "',";
                    sQry = sQry + "gibudscr = '" + gibudscr + "',";
                    sQry = sQry + "ntamt = " + ntamt + ",";
                    sQry = sQry + "jamt = " + jamt + ",";
                    sQry = sQry + "amt =" + Amt;

                    sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY407_DataFind();

                }
                else
                {
                    // 신규
                    sQry = "INSERT INTO [p_seoygibuhis]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "rel,";
                    sQry = sQry + "kname,";
                    sQry = sQry + "juminno,";
                    sQry = sQry + "gibucd,";
                    sQry = sQry + "gibudscr,";
                    sQry = sQry + "saupno,";
                    sQry = sQry + "sangho,";
                    sQry = sQry + "ntamt,";
                    sQry = sQry + "jamt,";
                    sQry = sQry + "amt)";

                    sQry = sQry + " VALUES(";

                    sQry = sQry + "'" + saup + "',";
                    sQry = sQry + "'" + yyyy + "',";
                    sQry = sQry + "'" + sabun + "',";
                    sQry = sQry + "'" + rel + "',";
                    sQry = sQry + "'" + kname + "',";
                    sQry = sQry + "'" + juminno + "',";
                    sQry = sQry + "'" + gibucd + "',";
                    sQry = sQry + "'" + gibudscr + "',";
                    sQry = sQry + "'" + saupno + "',";
                    sQry = sQry + "'" + sangho + "',";
                    sQry = sQry + ntamt + ",";
                    sQry = sQry + jamt + ",";
                    sQry = sQry + Amt + " )";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY407_DataFind();
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 0)
                { }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_Delete
        /// </summary>
        private void PH_PY407_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string saup, yyyy, sabun, gibucd, saupno, juminno, FullName = string.Empty;
            double cnt = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.VALUE.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.VALUE.ToString().Trim();
                gibucd = oForm.Items.Item("gibucd").Specific.VALUE.ToString().Trim();
                saupno = oForm.Items.Item("saupno").Specific.VALUE.ToString().Trim();
                FullName = oForm.Items.Item("kname").Specific.VALUE.ToString().Trim();

                sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

                oRecordSet.DoQuery(sQry);

                cnt = oRecordSet.Fields.Item(0).Value;
                if (cnt > 0)
                {

                    if (PSH_Globals.SBO_Application.MessageBox(" 선택한대상자('" + FullName + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                    {
                        sQry = "Delete From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                        sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY407_DataFind();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_TitleSetting
        /// </summary>
        private void PH_PY407_TitleSetting(ref int iRow)
        {
            int i = 0;
            string[] COLNAM = new string[15];

            try
            {
                COLNAM[0] = "연도";
                COLNAM[1] = "관계";
                COLNAM[2] = "관계명";
                COLNAM[3] = "성명";
                COLNAM[4] = "주민번호";
                COLNAM[5] = "기부금코드";
                COLNAM[6] = "기부금명";
                COLNAM[7] = "기부내용";
                COLNAM[8] = "사업자번호";
                COLNAM[9] = "기부처명";
                COLNAM[10] = "기부금액(국세청)";
                COLNAM[11] = "기부금액(국세청외)";
                COLNAM[12] = "기부장려금신청금액";
                COLNAM[13] = "사번";
                COLNAM[14] = "사업장";

                for (i = 0; i <= Information.UBound(COLNAM); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    oGrid1.Columns.Item(i).Editable = false;
                    //if (COLNAM[i] == "사번" | COLNAM[i] == "공제구분코드" | COLNAM[i] == "공제대상코드" | COLNAM[i] == "관계코드" | COLNAM[i] == "주민번호")
                    //{
                    //    oGrid1.Columns.Item(i).Visible = false;
                    //}
                    
                }
                oGrid1.Columns.Item(10).RightJustified = true;
                oGrid1.Columns.Item(11).RightJustified = true;
                oGrid1.Columns.Item(12).RightJustified = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
//	internal class PH_PY407
//	{
//////********************************************************************************
//////  File           : PH_PY407.cls
//////  Module         : 인사관리 > 연말정산관리
//////  Desc           : 정산기부금등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Grid oGrid1;
//		public SAPbouiCOM.Matrix oMat;
//		public SAPbouiCOM.DataTable oDS_PH_PY407A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY407L;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY407.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY407_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY407");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "Code"

//			oForm.PaneLevel = 1;
//			oForm.Freeze(true);
//			PH_PY407_CreateItems();
//			PH_PY407_FormItemEnabled();
//			PH_PY407_EnableMenus();
//			//    Call PH_PY407_SetDocument(oFromDocEntry01)
//			//    Call PH_PY407_FormResize

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

//		private bool PH_PY407_CreateItems()
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

//			oDS_PH_PY407L = oForm.DataSources.DBDataSources("@PS_USERDS01");

//			oGrid1 = oForm.Items.Item("Grid01").Specific;

//			oMat = oForm.Items.Item("Mat01").Specific;
//			oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

//			oForm.DataSources.DataTables.Add("PH_PY407");

//			oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY407");
//			oDS_PH_PY407A = oForm.DataSources.DataTables.Item("PH_PY407");


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			////사업장
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

//			////이름
//			oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

//			////관계
//			oCombo = oForm.Items.Item("rel").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("1", "거주자");
//			oCombo.ValidValues.Add("2", "배우자");
//			oCombo.ValidValues.Add("3", "직계비속");
//			oCombo.ValidValues.Add("4", "직계존속");
//			oCombo.ValidValues.Add("5", "형제,자매");
//			oCombo.ValidValues.Add("6", "그외");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("rel").DisplayDesc = true;

//			////성명
//			oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

//			////주민번호
//			oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

//			////기부금코드  73
//			oCombo = oForm.Items.Item("gibucd").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");

//			////기부내용 2018추가
//			oCombo = oForm.Items.Item("gibudscr").Specific;
//			oCombo.ValidValues.Add("1", "금전");
//			oCombo.ValidValues.Add("2", "현물");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//			oForm.Items.Item("gibudscr").DisplayDesc = true;

//			////사업자번호
//			oForm.DataSources.UserDataSources.Add("saupno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("saupno").Specific.DataBind.SetBound(true, "", "saupno");

//			////기부처명
//			oForm.DataSources.UserDataSources.Add("sangho", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("sangho").Specific.DataBind.SetBound(true, "", "sangho");

//			////공제금액(국세청)
//			oForm.DataSources.UserDataSources.Add("ntamt", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ntamt").Specific.DataBind.SetBound(true, "", "ntamt");

//			////공제금액(국세청외)
//			oForm.DataSources.UserDataSources.Add("amt", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("amt").Specific.DataBind.SetBound(true, "", "amt");

//			////기부장려금신청금액 2016
//			oForm.DataSources.UserDataSources.Add("jamt", SAPbouiCOM.BoDataType.dt_SUM);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("jamt").Specific.DataBind.SetBound(true, "", "jamt");


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
//			PH_PY407_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY407_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", false);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", false);
//			////행삭제

//			return;
//			PH_PY407_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY407_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY407_FormItemEnabled();
//				//        Call PH_PY407_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY407_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY407_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY407_FormItemEnabled()
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

//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE))) {
//					//UPGRADE_WARNING: oForm.Items(Year).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("Year").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;
//				}

//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE))) {
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("MSTCOD").Specific.VALUE = "";
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("FullName").Specific.VALUE = "";
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("TeamName").Specific.VALUE = "";
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("RspName").Specific.VALUE = "";
//					//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oForm.Items.Item("ClsName").Specific.VALUE = "";
//				}

//				oCombo = oForm.Items.Item("rel").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oForm.DataSources.UserDataSources.Item("kname").Value = "";
//				oForm.DataSources.UserDataSources.Item("juminno").Value = "";

//				oCombo = oForm.Items.Item("gibucd").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oCombo = oForm.Items.Item("gibudscr").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oForm.DataSources.UserDataSources.Item("saupno").Value = "";
//				oForm.DataSources.UserDataSources.Item("sangho").Value = "";

//				oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);


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

//			////Key set
//			oForm.Items.Item("CLTCOD").Enabled = true;
//			oForm.Items.Item("Year").Enabled = true;
//			oForm.Items.Item("MSTCOD").Enabled = true;

//			oForm.Items.Item("juminno").Enabled = true;
//			oForm.Items.Item("saupno").Enabled = true;
//			oForm.Items.Item("gibucd").Enabled = true;

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY407_FormItemEnabled_Error:

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//			string Div = null;
//			string target = null;
//			string relate = null;
//			string FullName = null;
//			string YY = null;
//			string Result = null;
//			string yyyy = null;


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
//							if (PH_PY407_DataValidCheck() == false) {
//								BubbleEvent = false;
//							}
//						}

//						if (pval.ItemUID == "Btn_ret") {
//							PH_PY407_MTX01();
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
//								PH_PY407_SAVE();
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
//								PH_PY407_Delete();
//								PH_PY407_FormItemEnabled();
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
//										PH_PY407_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY407_FormItemEnabled();
//									}
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									if (pval.ActionSuccess == true) {
//										PH_PY407_FormItemEnabled();
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

//							if (pval.ItemUID == "div") {
//								//UPGRADE_WARNING: oForm.Items(div).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("div").Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							}
//							if (pval.ItemUID == "target") {
//								//UPGRADE_WARNING: oForm.Items(target).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("target").Specific.VALUE)) {
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
//							if (pval.ItemUID == "rel") {
//								oMat.Clear();
//								oDS_PH_PY407L.Clear();

//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE;
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								relate = oForm.Items.Item("rel").Specific.VALUE;

//								sQry = "EXEC [PH_PY407_03] '" + MSTCOD + "', '" + relate + "'";

//								oRecordSet.DoQuery(sQry);

//								for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//									if (i + 1 > oDS_PH_PY407L.Size) {
//										oDS_PH_PY407L.InsertRecord((i));
//									}

//									oMat.AddRow();
//									oDS_PH_PY407L.Offset = i;

//									oDS_PH_PY407L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//									oDS_PH_PY407L.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet.Fields.Item("kname").Value));
//									oDS_PH_PY407L.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet.Fields.Item("juminno").Value));
//									oDS_PH_PY407L.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet.Fields.Item("birthymd").Value));
//									oDS_PH_PY407L.SetValue("U_ColReg04", i, Strings.Trim(oRecordSet.Fields.Item("relate").Value));
//									oDS_PH_PY407L.SetValue("U_ColReg05", i, Strings.Trim(oRecordSet.Fields.Item("addr").Value));
//									oRecordSet.MoveNext();
//								}

//								oMat.LoadFromDataSource();
//								oMat.AutoResizeColumns();

//								if ((oRecordSet.RecordCount == 0)) {
//									oForm.DataSources.UserDataSources.Item("kname").Value = "";
//									oForm.DataSources.UserDataSources.Item("juminno").Value = "";
//									oCombo = oForm.Items.Item("gibucd").Specific;
//									oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
//									oForm.DataSources.UserDataSources.Item("saupno").Value = "";
//									oForm.DataSources.UserDataSources.Item("sangho").Value = "";

//									oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
//									oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
//									oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);

//								}

//								if ((oRecordSet.RecordCount == 1)) {
//									//UPGRADE_WARNING: oForm.Items(kname).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oMat.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("kname").Specific.VALUE = oMat.Columns.Item("kname").Cells.Item(1).Specific.VALUE;
//									//UPGRADE_WARNING: oForm.Items(juminno).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oMat.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("juminno").Specific.VALUE = oMat.Columns.Item("juminno").Cells.Item(1).Specific.VALUE;
//								}


//							}

//						}
//					}

//					oForm.Freeze(false);
//					//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					oRecordSet = null;
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


//											//Call oMat.SelectRow(pval.Row, True, False)
//											PH_PY407_MTX02(pval.ItemUID, ref pval.Row, ref pval.ColUID);
//											break;
//									}

//								}
//								break;
//							case "Mat01":
//								if (pval.Row >= 0) {
//									oMat.SelectRow(pval.Row, true, false);
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
//						if (pval.ItemUID == "Mat01") {

//							//UPGRADE_WARNING: oForm.Items(kname).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oMat.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("kname").Specific.VALUE = oMat.Columns.Item("kname").Cells.Item(pval.Row).Specific.VALUE;
//							//UPGRADE_WARNING: oForm.Items(juminno).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							//UPGRADE_WARNING: oMat.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oForm.Items.Item("juminno").Specific.VALUE = oMat.Columns.Item("juminno").Cells.Item(pval.Row).Specific.VALUE;

//						}
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

//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
//									//                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
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
//								case "FullName":
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									FullName = oForm.Items.Item("FullName").Specific.VALUE;

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
//									sQry = sQry + " And U_status <> '5'";
//									//퇴사자 제외
//									sQry = sQry + " and U_FullName = '" + FullName + "'";

//									oRecordSet.DoQuery(sQry);

//									//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
//									//                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
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
//								case "ntamt":
//									break;

//								//                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
//								//                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
//								//                                    oForm.Items("ntsamt").Specific.VALUE = 0
//								//                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
//								//                                End If
//								//                            End If
//								//
//								case "juminno":
//									//주민번호
//									//주민번호입력시 생년월일 생성
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (Strings.Len(Strings.Trim(oForm.Items.Item("juminno").Specific.VALUE)) != 13) {
//										// oForm.Items("birthymd").Specific.VALUE = ""
//										MDC_Globals.Sbo_Application.MessageBox("주민번호자릿수가 틀립니다. 확인하세요.");
//									} else {
//										//                                If Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "1" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "2" Then
//										//                                    oForm.Items("birthymd").Specific.VALUE = "19" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
//										//                                ElseIf Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "3" Or Mid(oForm.Items("juminno").Specific.VALUE, 7, 1) = "4" Then
//										//                                    oForm.Items("birthymd").Specific.VALUE = "20" + Mid(oForm.Items("juminno").Specific.VALUE, 1, 6)
//										//                                End If
//									}
//									break;
//								case "amt":
//									break;
//								//                            If oForm.Items("handoamt").Specific.VALUE > 0 Then
//								//                                If CDbl(oForm.Items("ntsamt").Specific.VALUE) + CDbl(oForm.Items("amt").Specific.VALUE) > CDbl(oForm.Items("handoamt").Specific.VALUE) Then
//								//                                    oForm.Items("amt").Specific.VALUE = 0
//								//                                    Sbo_Application.MessageBox ("한도금액보다 초과됩니다. 확인하세요")
//								//                                End If
//								//                            End If
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
//						//                oMat.LoadFromDataSource
//						//                Call PH_PY407_AddMatrixRow

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
//						//UPGRADE_NOTE: oDS_PH_PY407A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY407A = null;

//						//                Set oMat = Nothing
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
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY407A", "Code")
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
//					//                Call PH_PY407_FormItemEnabled
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY407_FormItemEnabled();
//						break;
//					//                Call PH_PY407_AddMatrixRow
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY407_FormItemEnabled();
//						//                Call PH_PY407_AddMatrixRow
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY407_FormItemEnabled();
//						break;
//					//                Call PH_PY407_AddMatrixRow
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY407_FormItemEnabled();
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


//		public void PH_PY407_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY407'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY407_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY407_DataValidCheck()
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
//			PH_PY407_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY407_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;
//			int iRow = 0;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("MSTCOD").Specific.VALUE;

//			if (string.IsNullOrEmpty(Strings.Trim(Param01))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_MTX01_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Param02))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_MTX01_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(Param03))) {
//				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_MTX01_Exit;
//			}



//			sQry = "EXEC PH_PY407_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";

//			oDS_PH_PY407A.ExecuteQuery(sQry);



//			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

//			PH_PY407_TitleSetting(ref iRow);

//			oForm.Update();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY407_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY407_MTX01_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private void PH_PY407_MTX02(string oUID, ref int oRow = 0, ref string oCol = "")
//		{


//			////그리드 자료를 head에 로드

//			int i = 0;
//			string sQry = null;
//			int sRow = 0;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;
//			string Param05 = null;
//			string Param06 = null;

//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sRow = oRow;


//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oDS_PH_PY407A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oDS_PH_PY407A.Columns.Item("연도").Cells.Item(oRow).Value;
//			//UPGRADE_WARNING: oDS_PH_PY407A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oDS_PH_PY407A.Columns.Item("사번").Cells.Item(oRow).Value;
//			//UPGRADE_WARNING: oDS_PH_PY407A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oDS_PH_PY407A.Columns.Item("사업자번호").Cells.Item(oRow).Value;
//			//UPGRADE_WARNING: oDS_PH_PY407A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param05 = oDS_PH_PY407A.Columns.Item("기부금코드").Cells.Item(oRow).Value;
//			//UPGRADE_WARNING: oDS_PH_PY407A.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param06 = oDS_PH_PY407A.Columns.Item("주민번호").Cells.Item(oRow).Value;


//			sQry = "EXEC PH_PY407_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "'";
//			oRecordSet.DoQuery(sQry);

//			if ((oRecordSet.RecordCount == 0)) {

//				//  oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY407A.Columns.Item("MSTCOD").Cells(oRow).VALUE
//				//  oForm.Items("FullName").Specific.VALUE = oDS_PH_PY407A.Columns.Item("FullName").Cells(oRow).VALUE

//				oCombo = oForm.Items.Item("rel").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oForm.DataSources.UserDataSources.Item("kname").Value = "";
//				oForm.DataSources.UserDataSources.Item("juminno").Value = "";

//				oCombo = oForm.Items.Item("gibucd").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oCombo = oForm.Items.Item("gibudscr").Specific;
//				oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

//				oForm.DataSources.UserDataSources.Item("saupno").Value = "";
//				oForm.DataSources.UserDataSources.Item("sangho").Value = "";

//				oForm.DataSources.UserDataSources.Item("ntamt").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("amt").Value = Convert.ToString(0);
//				oForm.DataSources.UserDataSources.Item("jamt").Value = Convert.ToString(0);

//				//oForm.Items("TeamName").Specific.VALUE = ""
//				//oForm.Items("RspName").Specific.VALUE = ""
//				//oForm.Items("ClsName").Specific.VALUE = ""

//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY407_MTX02_Exit;
//			}

//			//oForm.Items("Year").Specific.VALUE = oRecordSet.Fields("Year").VALUE
//			//oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
//			//oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE

//			//    '//부서
//			//oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
//			//oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
//			//oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE

//			oCombo = oForm.Items.Item("gibucd").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("gibucd").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("gibudscr").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("gibudscr").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			oCombo = oForm.Items.Item("rel").Specific;
//			oCombo.Select(oRecordSet.Fields.Item("rel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("saupno").Value = oRecordSet.Fields.Item("saupno").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("sangho").Value = oRecordSet.Fields.Item("sangho").Value;

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("ntamt").Value = oRecordSet.Fields.Item("ntamt").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("amt").Value = oRecordSet.Fields.Item("amt").Value;
//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.DataSources.UserDataSources.Item("jamt").Value = oRecordSet.Fields.Item("jamt").Value;

//			oForm.Update();

//			oForm.ActiveItem = "rel";

//			////key set
//			oForm.Items.Item("CLTCOD").Enabled = false;
//			oForm.Items.Item("Year").Enabled = false;
//			oForm.Items.Item("MSTCOD").Enabled = false;


//			oForm.Items.Item("gibucd").Enabled = false;
//			oForm.Items.Item("juminno").Enabled = false;
//			oForm.Items.Item("saupno").Enabled = false;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY407_MTX02_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY407_MTX02_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY407_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY407A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY407A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY407_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY407_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY407_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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


//		private void PH_PY407_SAVE()
//		{

//			////데이타 저장

//			int i = 0;
//			string sQry = null;

//			string FullName = null;
//			string saup = null;
//			string sabun = null;
//			string yyyy = null;
//			string saupno = null;
//			string juminno = null;
//			string rel = null;
//			string kname = null;
//			string gibucd = null;
//			string sangho = null;
//			object gibudscr = null;

//			object Amt = null;
//			object ntamt = null;
//			double jamt = 0;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sabun = oForm.Items.Item("MSTCOD").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			rel = oForm.Items.Item("rel").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			kname = oForm.Items.Item("kname").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			juminno = oForm.Items.Item("juminno").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			gibucd = oForm.Items.Item("gibucd").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: gibudscr 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			gibudscr = oForm.Items.Item("gibudscr").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saupno = oForm.Items.Item("saupno").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sangho = oForm.Items.Item("sangho").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ntamt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ntamt = oForm.Items.Item("ntamt").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: Amt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Amt = oForm.Items.Item("amt").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			jamt = oForm.Items.Item("jamt").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FullName = oForm.Items.Item("FullName").Specific.VALUE;

//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			if (string.IsNullOrEmpty(Strings.Trim(yyyy))) {
//				MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_SAVE_Exit;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(saup))) {
//				MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_SAVE_Exit;
//			}
//			if (string.IsNullOrEmpty(Strings.Trim(sabun))) {
//				MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_SAVE_Exit;
//			}

//			//UPGRADE_WARNING: Amt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: ntamt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(juminno)) | (ntamt == 0 & Amt == 0)) {
//				MDC_Com.MDC_GF_Message(ref "정상적인 자료가 아닙니다. 확인바랍니다..", ref "E");
//				goto PH_PY407_SAVE_Exit;
//			}

//			sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//			sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.Fields.Item(0).Value > 0) {
//				////갱신
//				sQry = "Update [p_seoygibuhis] set ";
//				sQry = sQry + "rel = '" + rel + "',";
//				sQry = sQry + "kname = '" + kname + "',";
//				sQry = sQry + "sangho = '" + sangho + "',";
//				//UPGRADE_WARNING: gibudscr 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "gibudscr = '" + gibudscr + "',";
//				//UPGRADE_WARNING: ntamt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "ntamt = " + ntamt + ",";
//				sQry = sQry + "jamt = " + jamt + ",";
//				//UPGRADE_WARNING: Amt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "amt =" + Amt;

//				sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//				sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

//				oRecordSet.DoQuery(sQry);

//			} else {

//				////신규
//				sQry = "INSERT INTO [p_seoygibuhis]";
//				sQry = sQry + " (";
//				sQry = sQry + "saup,";
//				sQry = sQry + "yyyy,";
//				sQry = sQry + "sabun,";
//				sQry = sQry + "rel,";
//				sQry = sQry + "kname,";
//				sQry = sQry + "juminno,";
//				sQry = sQry + "gibucd,";
//				sQry = sQry + "gibudscr,";
//				sQry = sQry + "saupno,";
//				sQry = sQry + "sangho,";
//				sQry = sQry + "ntamt,";
//				sQry = sQry + "jamt,";
//				sQry = sQry + "amt)";

//				sQry = sQry + " VALUES(";

//				sQry = sQry + "'" + saup + "',";
//				sQry = sQry + "'" + yyyy + "',";
//				sQry = sQry + "'" + sabun + "',";
//				sQry = sQry + "'" + rel + "',";
//				sQry = sQry + "'" + kname + "',";
//				sQry = sQry + "'" + juminno + "',";
//				sQry = sQry + "'" + gibucd + "',";
//				//UPGRADE_WARNING: gibudscr 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + "'" + gibudscr + "',";
//				sQry = sQry + "'" + saupno + "',";
//				sQry = sQry + "'" + sangho + "',";

//				//UPGRADE_WARNING: ntamt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + ntamt + ",";
//				sQry = sQry + jamt + ",";
//				//UPGRADE_WARNING: Amt 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + Amt + " )";

//				oRecordSet.DoQuery(sQry);
//			}


//			PH_PY407_FormItemEnabled();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			PH_PY407_MTX01();

//			return;
//			PH_PY407_SAVE_Exit:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);

//			return;
//			PH_PY407_SAVE_Error:
//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_SAVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY407_Delete()
//		{
//			////선택된 자료 삭제
//			int i = 0;
//			string sQry = null;
//			short cnt = 0;

//			string FullName = null;
//			string saup = null;
//			string sabun = null;
//			string yyyy = null;
//			string saupno = null;
//			string juminno = null;
//			string rel = null;
//			string kname = null;
//			string gibucd = null;
//			string sangho = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saup = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			yyyy = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sabun = oForm.Items.Item("MSTCOD").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			juminno = oForm.Items.Item("juminno").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			gibucd = oForm.Items.Item("gibucd").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			saupno = oForm.Items.Item("saupno").Specific.VALUE;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FullName = oForm.Items.Item("kname").Specific.VALUE;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//			sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			cnt = oRecordSet.Fields.Item(0).Value;
//			if (cnt > 0) {

//				if (string.IsNullOrEmpty(Strings.Trim(yyyy))) {
//					MDC_Com.MDC_GF_Message(ref "년도가 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY407_Delete_Exit;
//				}

//				if (string.IsNullOrEmpty(Strings.Trim(saup))) {
//					MDC_Com.MDC_GF_Message(ref "사업장이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY407_Delete_Exit;
//				}
//				if (string.IsNullOrEmpty(Strings.Trim(sabun))) {
//					MDC_Com.MDC_GF_Message(ref "사번이 없습니다. 확인바랍니다..", ref "E");
//					goto PH_PY407_Delete_Exit;
//				}


//				if (MDC_Globals.Sbo_Application.MessageBox(" 선택한대상자('" + FullName + "')을 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1")) {
//					sQry = "Delete From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
//					sQry = sQry + " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
//					oRecordSet.DoQuery(sQry);
//				}
//			}


//			oForm.Freeze(false);


//			PH_PY407_MTX01();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;


//			return;
//			PH_PY407_Delete_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			return;
//			PH_PY407_Delete_Error:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY407_Delete_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY407_TitleSetting(ref int iRow)
//		{
//			//    Dim i               As Long
//			//    Dim j               As Long
//			//    Dim sQry            As String
//			//
//			//    Dim COLNAM(12)       As String
//			//
//			//    Dim oColumn         As SAPbouiCOM.EditTextColumn
//			//    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
//			//
//			//    Dim oRecordSet  As SAPbobsCOM.Recordset
//			//
//			//    On Error GoTo Error_Message
//			//
//			//    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
//			//
//			//    oForm.Freeze True
//			//
//			//    COLNAM(0) = "년도"
//			//    COLNAM(1) = "사번"
//			//    COLNAM(2) = "공제구분코드"
//			//    COLNAM(3) = "공제구분"
//			//    COLNAM(4) = "공제대상코드"
//			//    COLNAM(5) = "공제대상"
//			//    COLNAM(6) = "관계코드"
//			//    COLNAM(7) = "관계"
//			//    COLNAM(8) = "대상자성명"
//			//    COLNAM(9) = "주민번호"
//			//    COLNAM(10) = "금액(국세청)"
//			//    COLNAM(11) = "금액(국세청외)"
//			//    COLNAM(12) = "합계금액"
//			//
//			//    For i = 0 To UBound(COLNAM)
//			//        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
//			//        oGrid1.Columns.Item(i).Editable = False
//			//        If COLNAM(i) = "사번" Or COLNAM(i) = "공제구분코드" Or COLNAM(i) = "공제대상코드" Or COLNAM(i) = "관계코드" Or COLNAM(i) = "주민번호" Then
//			//            oGrid1.Columns.Item(i).Visible = False
//			//        End If
//			//
//			//        oGrid1.Columns.Item(i).RightJustified = True
//			//
//			//    Next i
//			//
//			//    oGrid1.AutoResizeColumns
//			//
//			//    oForm.Freeze False
//			//
//			//    Set oColumn = Nothing
//			//
//			//    Exit Sub
//			//
//			//Error_Message:
//			//    oForm.Freeze False
//			//    Set oColumn = Nothing
//			//    Sbo_Application.SetStatusBarMessage "PH_PY407_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
//		}
//	}
//}
