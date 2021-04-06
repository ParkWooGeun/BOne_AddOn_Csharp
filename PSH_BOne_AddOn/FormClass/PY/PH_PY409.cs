using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기부금조정명세자료등록
    /// </summary>
    internal class PH_PY409 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.Grid oGrid2;
        public SAPbouiCOM.DataTable oDS_PH_PY409;
        public SAPbouiCOM.DataTable oDS_PH_PY4091;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY409.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY409_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY409");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY409_CreateItems();
                PH_PY409_FormItemEnabled();
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
        private void PH_PY409_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                // Grid1
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY409");
                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY409");
                oDS_PH_PY409 = oForm.DataSources.DataTables.Item("PH_PY409");
                // Grid2
                oGrid2 = oForm.Items.Item("Grid02").Specific;
                oForm.DataSources.DataTables.Add("PH_PY4091");
                oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY4091");
                oDS_PH_PY4091 = oForm.DataSources.DataTables.Item("PH_PY4091");

                // 그리드1 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("연도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("기부금코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("기부금명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("기부년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("기부금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("전년까지공제된금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("공제대상금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("해당년도공제금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("소멸금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("이월금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY409").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 그리드2 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("기부자구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("총계", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("법정(10) 5Y", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("정치자금(20) X", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("특례(30) 2Y", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("공익법인신탁(31) 3Y", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("종교단체외지정(40) 5Y", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("종교단체지정(41) 5Y", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("우리사주조합(42) X", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY4091").Columns.Add("공제제외", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
                oForm.DataSources.UserDataSources.Item("Year").Value = DateTime.Now.AddYears(-1).ToString("yyyy");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 이름(조회)
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 부서명(조회)
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명(조회)
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명(조회)
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 이월(조회)
                oForm.DataSources.UserDataSources.Add("dontew", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dontew").Specific.DataBind.SetBound(true, "", "dontew");

                // 세액공제대상(조회)
                oForm.DataSources.UserDataSources.Add("gongamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("gongamt").Specific.DataBind.SetBound(true, "", "gongamt");

                // 합계(조회)
                oForm.DataSources.UserDataSources.Add("donttot", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("donttot").Specific.DataBind.SetBound(true, "", "donttot");

                // 기부금코드
                oForm.DataSources.UserDataSources.Add("gcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("gcode").Specific, "Y");
                oForm.Items.Item("gcode").DisplayDesc = true;

                // 기부년도
                oForm.DataSources.UserDataSources.Add("gyyyy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("gyyyy").Specific.DataBind.SetBound(true, "", "gyyyy");

                // 기부금액
                oForm.DataSources.UserDataSources.Add("gibuamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("gibuamt").Specific.DataBind.SetBound(true, "", "gibuamt");

                // 전년까지공제된금액
                oForm.DataSources.UserDataSources.Add("jgamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("jgamt").Specific.DataBind.SetBound(true, "", "jgamt");

                // 공제대상금액
                oForm.DataSources.UserDataSources.Add("gamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("gamt").Specific.DataBind.SetBound(true, "", "gamt");

                // 해당년도공제금액
                oForm.DataSources.UserDataSources.Add("ygamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ygamt").Specific.DataBind.SetBound(true, "", "ygamt");

                // 소멸금액
                oForm.DataSources.UserDataSources.Add("disamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("disamt").Specific.DataBind.SetBound(true, "", "disamt");

                // 이월금액
                oForm.DataSources.UserDataSources.Add("ewamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ewamt").Specific.DataBind.SetBound(true, "", "ewamt");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY409_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY409_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                
                if (oForm.Items.Item("Year").Specific.Value == "")
                {
                    oForm.DataSources.UserDataSources.Item("Year").Value = DateTime.Now.AddYears(-1).ToString("yyyy");
                }
                if (oForm.Items.Item("MSTCOD").Specific.Value == "")
                {
                    oForm.Items.Item("MSTCOD").Specific.Value = "";
                    oForm.Items.Item("FullName").Specific.Value = "";
                    oForm.Items.Item("TeamName").Specific.Value = "";
                    oForm.Items.Item("RspName").Specific.Value = "";
                    oForm.Items.Item("ClsName").Specific.Value = "";
                    oForm.Items.Item("gongamt").Specific.Value = "0";
                }

                oForm.Items.Item("gcode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("gyyyy").Value = "";
                oForm.DataSources.UserDataSources.Item("gibuamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("jgamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("gamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("ygamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("disamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("ewamt").Value = "0";

                // Key set
                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;
                oForm.Items.Item("gcode").Enabled = true;
                oForm.Items.Item("gyyyy").Enabled = true;

                // 문서추가
                oForm.EnableMenu("1282", true); 

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY409_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
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
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            break;
                        case "1282": //문서추가
                            PH_PY409_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "1293": //행삭제
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY409);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY4091);
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
            string Year = string.Empty;
            string YN = string.Empty;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY409_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                        sQry = "select UseYN = b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + Year + "'";
                        oRecordSet.DoQuery(sQry);

                        YN = oRecordSet.Fields.Item("UseYN").Value.ToString().Trim();
                        if (YN != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (YN == "Y")
                        {
                            PH_PY409_SAVE();
                            PH_PY409_DataFind();
                            PH_PY409_FormItemEnabled();
                        }
                        
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                        sQry = "select UseYN = b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + Year + "'";
                        oRecordSet.DoQuery(sQry);

                        YN = oRecordSet.Fields.Item("UseYN").Value.ToString().Trim();
                        if (YN != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (YN == "Y")
                        {
                            PH_PY409_Delete();
                            PH_PY409_DataFind();
                            PH_PY409_FormItemEnabled();
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            //string sQry = string.Empty;
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
            string Year = string.Empty;
            string MSTCOD = string.Empty;
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
                            case "Year":
                                oForm.Items.Item("MSTCOD").Specific.Value = "";
                                break;

                            case "MSTCOD":

                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

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
    ;                           sQry = sQry + " And U_Char3 = U_CLTCOD),'')";
                                sQry = sQry + " From [@PH_PY001A]";
                                sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry = sQry + " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value.ToString().Trim();
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value.ToString().Trim();

                                // 정산금액 찿기
                                sQry = "Select dont = dont_t + dont1_t + dont2_t + dont3_t + Isnull(poldont_t,0), ";
                                sQry = sQry + " dontew = dontew, ";
                                sQry = sQry + " donttt = dont_t + dont1_t + dont2_t + dont3_t + Isnull(poldont_t,0) + dontew ";
                                sQry = sQry + " From [p_seoycpt] ";
                                sQry = sQry + " WHERE saup = '" + CLTCOD + "'";
                                sQry = sQry + " and yyyy = '" + Year + "'";
                                sQry = sQry + " and sabun = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("dontew").Specific.Value = oRecordSet.Fields.Item("dontew").Value.ToString().Trim();
                                oForm.Items.Item("gongamt").Specific.Value = oRecordSet.Fields.Item("dont").Value.ToString().Trim();
                                oForm.Items.Item("donttot").Specific.Value = oRecordSet.Fields.Item("donttt").Value.ToString().Trim();

                                // 기부자료집계 표시 Grid 2 
                                sQry = "EXEC PH_PY409_03 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                                oDS_PH_PY4091.ExecuteQuery(sQry); 
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
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            string Param04 = string.Empty;
            string Param05 = string.Empty;

            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);
                            Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            Param02 = oDS_PH_PY409.Columns.Item("연도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY409.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY409.Columns.Item("기부금코드").Cells.Item(pVal.Row).Value;
                            Param05 = oDS_PH_PY409.Columns.Item("기부년도").Cells.Item(pVal.Row).Value;

                            if (string.IsNullOrEmpty(Param02))
                            {
                                oForm.Items.Item("gcode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("gyyyy").Value = "";
                                oForm.DataSources.UserDataSources.Item("gibuamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("jgamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("gamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ygamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("disamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("ewamt").Value = "0";

                                oForm.Update();

                            }
                            else
                            {
                                sQry = "EXEC PH_PY409_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "'";
                                oRecordSet.DoQuery(sQry);

                                if ((oRecordSet.RecordCount == 0))
                                {
                                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                                }
                                else
                                {
                                    oForm.Items.Item("gcode").Specific.Select(oRecordSet.Fields.Item("gcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.DataSources.UserDataSources.Item("gyyyy").Value = oRecordSet.Fields.Item("gyyyy").Value;
                                    oForm.DataSources.UserDataSources.Item("gibuamt").Value = oRecordSet.Fields.Item("gibuamt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("jgamt").Value = oRecordSet.Fields.Item("jgamt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("ygamt").Value = oRecordSet.Fields.Item("ygamt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("disamt").Value = oRecordSet.Fields.Item("disamt").Value.ToString().Trim();
                                    oForm.DataSources.UserDataSources.Item("ewamt").Value = oRecordSet.Fields.Item("ewamt").Value.ToString().Trim();

                                    oForm.Update();
                                    oForm.ActiveItem = "gibuamt";

                                    //key set
                                    oForm.Items.Item("CLTCOD").Enabled = false;
                                    oForm.Items.Item("Year").Enabled = false;
                                    oForm.Items.Item("MSTCOD").Enabled = false;
                                    oForm.Items.Item("gcode").Enabled = false;
                                    oForm.Items.Item("gyyyy").Enabled = false;
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
        /// PH_PY409_DataFind
        /// </summary>
        private void PH_PY409_DataFind()
        {
            short ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string Year = string.Empty;
            string MSTCOD = string.Empty;

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(CLTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(MSTCOD))
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                sQry = "EXEC PH_PY409_01 '" + CLTCOD + "', '" + Year + "', '" + MSTCOD + "'";
                oDS_PH_PY409.ExecuteQuery(sQry);
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장을 입력 하세요, 확인바랍니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("사원코드를 입력 하세요, 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY409_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY409_SAVE
        /// </summary>
        private void PH_PY409_SAVE()
        {
            // 데이타 저장
            short ErrNum = 0;
            string sQry = string.Empty;
            string saup = string.Empty;
            string sabun = string.Empty;
            string FullName = string.Empty;
            string yyyy = string.Empty;
            string gcode = string.Empty;
            string gyyyy = string.Empty; ;
            double gibuamt = 0;
            double jgamt = 0;
            double gamt = 0;
            double ygamt = 0;
            double disamt = 0;
            double ewamt = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
               // oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                gcode = oForm.Items.Item("gcode").Specific.Value.ToString().Trim();
                gyyyy = oForm.Items.Item("gyyyy").Specific.Value.ToString().Trim();
                gibuamt = Convert.ToDouble(oForm.Items.Item("gibuamt").Specific.Value);
                jgamt = Convert.ToDouble(oForm.Items.Item("jgamt").Specific.Value);
                gamt = Convert.ToDouble(oForm.Items.Item("gamt").Specific.Value);
                ygamt = Convert.ToDouble(oForm.Items.Item("ygamt").Specific.Value);
                disamt = Convert.ToDouble(oForm.Items.Item("disamt").Specific.Value);
                ewamt = Convert.ToDouble(oForm.Items.Item("ewamt").Specific.Value);
                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();

                if (string.IsNullOrWhiteSpace(yyyy))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(saup))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(sabun))
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                sQry = " Select Count(*) From [p_seoygibucont] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                sQry = sQry + " And gcode = '" + gcode + "' And gyyyy = '" + gyyyy + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value <= 0)
                {
                    //신규
                    sQry = "INSERT INTO [p_seoygibucont]";
                    sQry = sQry + " (";
                    sQry = sQry + "saup,";
                    sQry = sQry + "yyyy,";
                    sQry = sQry + "sabun,";
                    sQry = sQry + "gcode,";
                    sQry = sQry + "gyyyy,";
                    sQry = sQry + "gibuamt,";
                    sQry = sQry + "jgamt,";
                    sQry = sQry + "gamt,";
                    sQry = sQry + "ygamt,";
                    sQry = sQry + "disamt,";
                    sQry = sQry + "ewamt)";

                    sQry = sQry + " VALUES(";
                    sQry = sQry + "'" + saup + "',";
                    sQry = sQry + "'" + yyyy + "',";
                    sQry = sQry + "'" + sabun + "',";
                    sQry = sQry + "'" + gcode + "',";
                    sQry = sQry + "'" + gyyyy + "',";

                    sQry = sQry + gibuamt + ",";
                    sQry = sQry + jgamt + ",";
                    sQry = sQry + gamt + ",";
                    sQry = sQry + ygamt + ",";
                    sQry = sQry + disamt + ",";
                    sQry = sQry + ewamt + " )";

                    oRecordSet.DoQuery(sQry);
                    // oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                else
                {
                    //수정
                    sQry = "Update [p_seoygibucont] set ";
                    sQry = sQry + "gibuamt = " + gibuamt + ",";
                    sQry = sQry + "jgamt = " + jgamt + ",";
                    sQry = sQry + "gamt = " + gamt + ",";
                    sQry = sQry + "ygamt = " + ygamt + ",";
                    sQry = sQry + "disamt = " + disamt + ",";
                    sQry = sQry + "ewamt = " + ewamt + "";
                    sQry = sQry + " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    sQry = sQry + " And gcode = '" + gcode + "' And gyyyy = '" + gyyyy + "'";
                    oRecordSet.DoQuery(sQry);
                    // oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY409_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
             //   oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY409_Delete
        /// </summary>
        private void PH_PY409_Delete()
        {
            // 데이타 삭제
            short ErrNum = 0;
            string sQry = string.Empty;
            string saup = string.Empty;
            string yyyy = string.Empty;
            string sabun = string.Empty;
            string gcode = string.Empty;
            string gyyyy = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //oForm.Freeze(true);

                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                gcode = oForm.Items.Item("gcode").Specific.Value.ToString().Trim();
                gyyyy = oForm.Items.Item("gyyyy").Specific.Value.ToString().Trim();

                if (string.IsNullOrWhiteSpace(yyyy))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(saup))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrWhiteSpace(sabun))
                {
                    ErrNum = 3;
                    throw new Exception();
                }

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", Convert.ToInt32("2"), "예", "아니오") == Convert.ToDouble("1"))
                {
                    if (oDS_PH_PY409.Rows.Count > 0)
                    {
                        sQry = "Delete From [p_seoygibucont] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                        sQry = sQry + " And gcode = '" + gcode + "' And gyyyy = '" + gyyyy + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY409_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                //oForm.Freeze(false);
            }
        }
    }
}
