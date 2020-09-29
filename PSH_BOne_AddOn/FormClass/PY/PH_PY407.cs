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
