using System;
using System.Collections.Generic;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    internal class PH_PY718 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY718;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 생산완료금액대비O/T현황
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY718.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY718_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY718");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY718_CreateItems();
                PH_PY718_EnableMenus();
                PH_PY718_SetDocument(oFromDocEntry01);
                //PH_PY718_FormResize();
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
        private void PH_PY718_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY718");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY718");
                oDS_PH_PY718 = oForm.DataSources.DataTables.Item("PH_PY718");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 기준년월(Fr)
                oForm.DataSources.UserDataSources.Add("FrYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("FrYM").Specific.DataBind.SetBound(true, "", "FrYM");
                oForm.Items.Item("FrYM").Specific.String = DateTime.Now.ToString("yyyyMM");

                // 기준년월(To)
                oForm.DataSources.UserDataSources.Add("ToYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("ToYM").Specific.DataBind.SetBound(true, "", "ToYM");
                oForm.Items.Item("ToYM").Specific.String = DateTime.Now.ToString("yyyyMM");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY718_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY718_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY718_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY718_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY718_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck : 입력데이터의 Valid Check
        /// </summary>
        /// <returns></returns>
        private bool PH_PY718_DataValidCheck()
        {
            bool functionReturnValue = false;

            short ErrNum = 0;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ////사원명 찿아서 화면 표시 하기
                //sQry = "        SELECT      U_FullName";
                //sQry = sQry + " FROM        [@PH_PY001A]";
                //sQry = sQry + " WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                //oRecordSet.DoQuery(sQry);
                //oForm.Items.Item("MSTNAM").Specific.Value = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();

                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_DataValidCheck_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// DataFind : 자료 조회
        /// </summary>
        private void PH_PY718_DataFind()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string FrYM = string.Empty;
            string ToYM = string.Empty;
            string CntcCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            

            try
            {
                PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
                ToYM = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
                CntcCode = dataHelpClass.User_MSTCOD();

                sQry = "                EXEC [PH_PY718_01] ";
                sQry = sQry + "'" + CLTCOD + "',";              //사업장
                sQry = sQry + "'" + FrYM + "',";                //기준년월(Fr)
                sQry = sQry + "'" + ToYM + "',";                //기준년월(To)
                sQry = sQry + "'" + CntcCode + "'";             //조회자 사번

                oDS_PH_PY718.ExecuteQuery(sQry);

                oGrid1.Columns.Item(2).RightJustified = true;
                oGrid1.Columns.Item(4).RightJustified = true;
                oGrid1.Columns.Item(5).RightJustified = true;
                oGrid1.Columns.Item(6).RightJustified = true;
                oGrid1.Columns.Item(7).RightJustified = true;
                oGrid1.Columns.Item(8).RightJustified = true;
                oGrid1.Columns.Item(9).RightJustified = true;
                oGrid1.Columns.Item(10).RightJustified = true;
                oGrid1.Columns.Item(11).RightJustified = true;
                oGrid1.Columns.Item(12).RightJustified = true;
                oGrid1.Columns.Item(13).RightJustified = true;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (PH_PY718_DataValidCheck() == true)
                        {
                            PH_PY718_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        if (PH_PY718_DataValidCheck() == true)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY718_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    //                If oForm.Mode = fm_ADD_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                ElseIf oForm.Mode = fm_OK_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                End If

                    //                If oForm.Mode = fm_ADD_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                ElseIf oForm.Mode = fm_OK_MODE Then
                    //                    If pval.ActionSuccess = True Then
                    //                        Call PH_PY718_FormItemEnabled
                    //                    End If
                    //                End If
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
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
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    //switch (pVal.ItemUID)
                    //{
                    //    case "Grid01":
                    //        if (pVal.Row > 0)
                    //        {
                    //            //oGrid1.SelectRow(pval.Row, True, False);
                    //        }
                    //        break;
                    //}

                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sQry = string.Empty;

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
                            //case "MSTCOD": //사원명
                            //    sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                            //    oRecordSet.DoQuery(sQry);
                            //    oForm.Items.Item("MSTNAM").Specific.Value = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                            //    break;
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
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY718_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY718);
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
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
                            PH_PY718_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY718_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY718_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY718_FormItemEnabled();
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }

                switch (pVal.ItemUID)
                {
                    case "Grid01":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY718_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string FrYM = string.Empty;
            string ToYM = string.Empty;
            string CntcCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
            FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
            ToYM = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
            CntcCode = dataHelpClass.User_MSTCOD();

            try
            {
                           
                WinTitle = "[PH_PY718] 생산완료금액대비O/T현황";
                ReportName = "PH_PY718_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                //Formula
                //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                
                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@FrYM", FrYM));
                dataPackParameter.Add(new PSH_DataPackClass("@ToYM", ToYM));
                dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY718_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        #region Event 코드

        #region Raise_FormItemEvent
        //		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			string sQry = null;
        //			int i = 0;
        //			SAPbouiCOM.ComboBox oCombo = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1

        //					if (pval.BeforeAction == true) {

        //					} else if (pval.BeforeAction == false) {


        //					if (pval.BeforeAction == true) {

        //					} else if (pval.BeforeAction == false) {


        //					}
        //					break;

        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					break;

        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3

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

        //					}
        //					oForm.Freeze(false);
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					if (pval.BeforeAction == true) {
        //						
        //					} else if (pval.BeforeAction == false) {

        //					}
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
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
        //					oForm.Freeze(true);
        //					if (pval.BeforeAction == true) {
        //					} else if (pval.BeforeAction == false) {

        //					}
        //					oForm.Freeze(false);
        //					break;
        //				//----------------------------------------------------------
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					if (pval.BeforeAction == true) {
        //					} else if (pval.BeforeAction == false) {

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
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormMenuEvent
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			int i = 0;
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			oForm.Freeze(true);

        //			if ((pval.BeforeAction == true)) {

        //			} else if ((pval.BeforeAction == false)) {

        //			}
        //			oForm.Freeze(false);
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if ((BusinessObjectInfo.BeforeAction == true)) {
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
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
        //			return;
        //			Raise_FormDataEvent_Error:


        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{

        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}
        //			switch (pval.ItemUID) {
        //				case "Grid01":
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
        #endregion

        #endregion
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
//	internal class PH_PY718
//	{
////****************************************************************************************************************
//////  File : PH_PY718.cls
//////  Module : 인사관리 > 급여관리 > 급여리포트
//////  Desc : 생산완료금액대비O/T현황
//////  FormType : PH_PY718
//////  Create Date(Start) : 2014.08.19
//////  Create Date(End) : 2014.08.20
//////  Creator : Song Myoung gyu
//////  Modified Date :
//////  Modifier :
//////  Company : Poongsan Holdings
////****************************************************************************************************************

//		public string oFormUniqueID01;
//		public SAPbouiCOM.Form oForm;
//		public SAPbouiCOM.Grid oGrid01;

//			//클래스에서 선택한 마지막 아이템 Uid값
//		private string oLastItemUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private string oLastColUID01;
//			//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oLastColRow01;

////*******************************************************************
//// .srf 파일로부터 폼을 로드한다.
////*******************************************************************
//		public void LoadForm(string oFromDocEntry01 = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			string oInnerXml = null;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY718.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);

//			//매트릭스의 타이틀높이와 셀높이를 고정
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}

//			oFormUniqueID01 = "PH_PY718_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID01, "PH_PY718");
//			////폼추가
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			//폼 할당
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID01);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			////oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

//			oForm.Freeze(true);
//			PH_PY718_CreateItems();
//			PH_PY718_ComboBox_Setting();
//			PH_PY718_CF_ChooseFromList();
//			PH_PY718_EnableMenus();
//			PH_PY718_SetDocument(oFromDocEntry01);
//			PH_PY718_FormResize();

//			oForm.EnableMenu("1283", false);
//			//삭제
//			oForm.EnableMenu("1286", false);
//			//닫기
//			oForm.EnableMenu("1287", false);
//			//복제
//			oForm.EnableMenu("1285", false);
//			//복원
//			oForm.EnableMenu("1284", false);
//			//취소
//			oForm.EnableMenu("1293", false);
//			//행삭제
//			oForm.EnableMenu("1281", false);
//			oForm.EnableMenu("1282", true);

//			//UPGRADE_WARNING: oForm.Items(FrYM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FrYM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");
//			//UPGRADE_WARNING: oForm.Items(ToYM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ToYM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

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

//		public void PH_PY718_MTX01()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY718_MTX01()
//			//해당모듈 : PH_PY718
//			//기능 : 데이터 조회
//			//인수 : 없음
//			//반환값 : 없음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			string sQry = null;
//			short ErrNum = 0;

//			//    Dim RecordSet01 As SAPbobsCOM.Recordset
//			//    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

//			string CLTCOD = null;
//			//사업장
//			string FrYM = null;
//			//기준년월(Fr)
//			string ToYM = null;
//			//기준년월(To)
//			string CntcCode = null;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FrYM = Strings.Trim(oForm.Items.Item("FrYM").Specific.VALUE);
//			//기준년월(Fr)
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ToYM = Strings.Trim(oForm.Items.Item("ToYM").Specific.VALUE);
//			//기준년월(To)
//			CntcCode = MDC_PS_Common.User_MSTCOD();

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", 100, false);

//			oForm.Freeze(true);

//			sQry = "                EXEC [PH_PY718_01] ";
//			sQry = sQry + "'" + CLTCOD + "',";
//			//사업장
//			sQry = sQry + "'" + FrYM + "',";
//			//기준년월(Fr)
//			sQry = sQry + "'" + ToYM + "',";
//			//기준년월(To)
//			sQry = sQry + "'" + CntcCode + "'";
//			//조회자 사번

//			oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");
//			oGrid01.DataTable.Clear();
//			oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);

//			ProgBar01.Value = 100;
//			ProgBar01.Text = "조회중...!";
//			ProgBar01.Stop();

//			oGrid01.Columns.Item(2).RightJustified = true;
//			oGrid01.Columns.Item(4).RightJustified = true;
//			oGrid01.Columns.Item(5).RightJustified = true;
//			oGrid01.Columns.Item(6).RightJustified = true;
//			oGrid01.Columns.Item(7).RightJustified = true;
//			oGrid01.Columns.Item(8).RightJustified = true;
//			oGrid01.Columns.Item(9).RightJustified = true;
//			oGrid01.Columns.Item(10).RightJustified = true;
//			oGrid01.Columns.Item(11).RightJustified = true;
//			oGrid01.Columns.Item(12).RightJustified = true;
//			oGrid01.Columns.Item(13).RightJustified = true;
//			//    oGrid01.Columns(14).RightJustified = True
//			//    oGrid01.Columns(15).RightJustified = True

//			if (oGrid01.Rows.Count == 0) {
//				ErrNum = 1;
//				goto PH_PY718_MTX01_Error;
//			}

//			oGrid01.AutoResizeColumns();
//			oForm.Update();

//			oForm.Freeze(false);

//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//    Set RecordSet01 = Nothing
//			return;
//			PH_PY718_MTX01_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//    ProgBar01.Stop
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;
//			//    Set RecordSet01 = Nothing

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "조회 결과가 없습니다. 확인하세요.", ref "W");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY718_MTX01_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//		}

//		private bool PH_PY718_HeaderSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : PH_PY718_HeaderSpaceLineDel()
//			//해당모듈 : PH_PY718
//			//기능 : 필수입력사항 체크
//			//인수 : 없음
//			//반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short ErrNum = 0;
//			ErrNum = 0;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			switch (true) {
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrYM").Specific.VALUE)):
//					//기준년월(Fr)
//					ErrNum = 1;
//					goto PH_PY718_HeaderSpaceLineDel_Error;
//					break;
//				case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToYM").Specific.VALUE)):
//					//기준년월(To)
//					ErrNum = 2;
//					goto PH_PY718_HeaderSpaceLineDel_Error;
//					break;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY718_HeaderSpaceLineDel_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "기준년월(Fr)은 필수조회 조건입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("FrYM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "기준년월(To)은 필수조회 조건입니다. 확인하세요.", ref "E");
//				oForm.Items.Item("ToYM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				//    ElseIf ErrNum = 3 Then
//				//        MDC_Com.MDC_GF_Message "사원번호는 필수사항입니다. 확인하세요.", "E"
//				//        Call oForm.Items("MSTCOD").CLICK(ct_Regular)
//				//    ElseIf ErrNum = 4 Then
//				//        MDC_Com.MDC_GF_Message "시작일자는 필수사항입니다. 확인하세요.", "E"
//				//        Call oForm.Items("FrDate").CLICK(ct_Regular)
//				//    ElseIf ErrNum = 5 Then
//				//        MDC_Com.MDC_GF_Message "시작시각은 필수사항입니다. 확인하세요.", "E"
//				//        Call oForm.Items("FrTime").CLICK(ct_Regular)
//				//    ElseIf ErrNum = 6 Then
//				//        MDC_Com.MDC_GF_Message "종료일자는 필수사항입니다. 확인하세요.", "E"
//				//        Call oForm.Items("FrDate").CLICK(ct_Regular)
//				//    ElseIf ErrNum = 7 Then
//				//        MDC_Com.MDC_GF_Message "종료시각은 필수사항입니다. 확인하세요.", "E"
//				//        Call oForm.Items("FrTime").CLICK(ct_Regular)
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

///// 메트릭스 필수 사항 check
//		private bool PH_PY718_MatrixSpaceLineDel()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			short ErrNum = 0;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			string sQry = null;

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY718_MatrixSpaceLineDel_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "라인 데이터가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 2) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 사원코드가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 3) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 시간이 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 4) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 등록일자가 없습니다. 확인하세요.", ref "E");
//			} else if (ErrNum == 5) {
//				MDC_Com.MDC_GF_Message(ref "" + i + 1 + "번 라인의 비가동코드가 없습니다. 확인하세요.", ref "E");
//			} else {
//				MDC_Com.MDC_GF_Message(ref "PH_PY718_MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void PH_PY718_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short i = 0;
//			short ErrNum = 0;
//			string sQry = null;
//			string ItemCode = null;

//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			switch (oUID) {

//				case "CLTCOD":
//					break;


//			}

//			oForm.Freeze(false);

//			return;
//			PH_PY718_FlushToItemValue_Error:

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "PH_PY718_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			}

//		}

/////폼의 아이템 사용지정
//		public void PH_PY718_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//				//        oForm.Items("ItemCode").Enabled = True

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				//        Call CLTCOD_Select(oForm, "SCLTCOD")

//			}

//			return;
//			PH_PY718_FormItemEnabled_Error:

//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			MDC_Com.MDC_GF_Message(ref "PH_PY718_FormItemEnabled_Error:" + Err().Number + " - " + Err().Description, ref "E");
//		}

/////아이템 변경 이벤트
//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;
//				////et_FORM_ACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;
//				////et_FORM_DEACTIVATE
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////20
//					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;
//				////et_LOST_FOCUS
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
//					break;
//			}
//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			////BeforeAction = True
//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					case "1282":
//						//추가
//						///추가버튼 클릭시 메트릭스 insertrow

//						//                Call PH_PY718_FormReset

//						//                oMat01.Clear
//						//                oMat01.FlushToDataSource
//						//                oMat01.LoadFromDataSource

//						//                oForm.Mode = fm_ADD_MODE
//						//                BubbleEvent = False
//						//                Call PH_PY718_LoadCaption

//						//oForm.Items("GCode").Click ct_Regular


//						return;

//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;

//					case "7169":
//						//엑셀 내보내기
//						break;

//				}
//			////BeforeAction = False
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1284":
//						//취소
//						break;
//					case "1286":
//						//닫기
//						break;
//					case "1293":
//						//행삭제
//						break;
//					case "1281":
//						//찾기
//						break;
//					////Call PH_PY718_FormItemEnabled '//UDO방식
//					case "1282":
//						//추가
//						break;
//					//                oMat01.Clear
//					//                oDS_PH_PY718A.Clear

//					//                Call PH_PY718_LoadCaption
//					//                Call PH_PY718_FormItemEnabled
//					////Call PH_PY718_FormItemEnabled '//UDO방식
//					////Call PH_PY718_AddMatrixRow(0, True) '//UDO방식
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						//레코드이동버튼
//						break;
//					////Call PH_PY718_FormItemEnabled

//					case "7169":
//						//엑셀 내보내기
//						break;

//				}
//			}
//			return;
//			Raise_FormMenuEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			////BeforeAction = True
//			if ((BusinessObjectInfo.BeforeAction == true)) {
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
//			////BeforeAction = False
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
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
//			return;
//			Raise_FormDataEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_RightClickEvent_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				if (pval.ItemUID == "PH_PY718") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}

//				///조회
//				if (pval.ItemUID == "BtnSearch") {

//					if (PH_PY718_HeaderSpaceLineDel() == false) {
//						BubbleEvent = false;
//						return;
//					}

//					PH_PY718_MTX01();

//				} else if (pval.ItemUID == "BtnPrint") {

//					if (PH_PY718_HeaderSpaceLineDel() == false) {
//						BubbleEvent = false;
//						return;
//					}

//					PH_PY718_Print_Report01();

//				}

//			} else if (pval.BeforeAction == false) {
//				if (pval.ItemUID == "PH_PY718") {
//					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//					} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//					}
//				}
//			}

//			return;
//			Raise_EVENT_ITEM_PRESSED_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "MSTCOD", "");
//				//사번

//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_KEY_DOWN_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}

//			return;
//			Raise_EVENT_CLICK_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//				PH_PY718_FlushToItemValue(pval.ItemUID);

//			}

//			return;
//			Raise_EVENT_COMBO_SELECT_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_DOUBLE_CLICK_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {

//			}
//			return;
//			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			if (pval.BeforeAction == true) {

//				if (pval.ItemChanged == true) {

//					PH_PY718_FlushToItemValue(pval.ItemUID);

//				}

//			} else if (pval.BeforeAction == false) {

//			}

//			oForm.Freeze(false);

//			return;
//			Raise_EVENT_VALIDATE_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY718_FormItemEnabled();
//				////Call PH_PY718_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
//			}
//			return;
//			Raise_EVENT_MATRIX_LOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				PH_PY718_FormResize();
//			}
//			return;
//			Raise_EVENT_RESIZE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {

//			} else if (pval.BeforeAction == false) {
//				//        If (pval.ItemUID = "ItemCode") Then
//				//            Dim oDataTable01 As SAPbouiCOM.DataTable
//				//            Set oDataTable01 = pval.SelectedObjects
//				//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
//				//            Set oDataTable01 = Nothing
//				//        End If
//				//        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
//				//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY718A", "U_CardCode,U_CardName")
//				//        End If
//			}
//			return;
//			Raise_EVENT_CHOOSE_FROM_LIST_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.ItemUID == "Mat01") {
//				if (pval.Row > 0) {
//					oLastItemUID01 = pval.ItemUID;
//					oLastColUID01 = pval.ColUID;
//					oLastColRow01 = pval.Row;
//				}
//			} else {
//				oLastItemUID01 = pval.ItemUID;
//				oLastColUID01 = "";
//				oLastColRow01 = 0;
//			}
//			return;
//			Raise_EVENT_GOT_FOCUS_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//				SubMain.RemoveForms(oFormUniqueID01);
//				//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oForm = null;
//				//UPGRADE_NOTE: oGrid01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oGrid01 = null;
//			}
//			return;
//			Raise_EVENT_FORM_UNLOAD_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			if ((oLastColRow01 > 0)) {
//				if (pval.BeforeAction == true) {

//				} else if (pval.BeforeAction == false) {

//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY718_CreateItems()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			string oQuery01 = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;
//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oGrid01 = oForm.Items.Item("Grid01").Specific;
//			oForm.DataSources.DataTables.Add("DataTable");
//			oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

//			//사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

//			//기준년월(Fr)
//			oForm.DataSources.UserDataSources.Add("FrYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("FrYM").Specific.DataBind.SetBound(true, "", "FrYM");

//			//기준년월(To)
//			oForm.DataSources.UserDataSources.Add("ToYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("ToYM").Specific.DataBind.SetBound(true, "", "ToYM");

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY718_CreateItems_Error:

//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

/////콤보박스 set
//		public void PH_PY718_ComboBox_Setting()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet01 = null;

//			oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet01 = null;

//			return;
//			PH_PY718_ComboBox_Setting_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_ComboBox_Setting_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY718_CF_ChooseFromList()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY718_CF_ChooseFromList_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY718_EnableMenus()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY718_EnableMenus_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY718_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY718_FormItemEnabled();
//				////Call PH_PY718_AddMatrixRow(0, True) '//UDO방식일때
//			} else {
//				//        oForm.Mode = fm_FIND_MODE
//				//        Call PH_PY718_FormItemEnabled
//				//        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
//				//        oForm.Items("1").Click ct_Regular
//			}
//			return;
//			PH_PY718_SetDocument_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY718_FormResize()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			return;
//			PH_PY718_FormResize_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY718_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY718_Print_Report01()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string WinTitle = null;
//			string ReportName = null;
//			string sQry = null;

//			string CLTCOD = null;
//			//사업장
//			string FrYM = null;
//			//기준년월(Fr)
//			string ToYM = null;
//			//기준년월(To)
//			string CntcCode = null;

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			FrYM = Strings.Trim(oForm.Items.Item("FrYM").Specific.VALUE);
//			//기준년월(Fr)
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			ToYM = Strings.Trim(oForm.Items.Item("ToYM").Specific.VALUE);
//			//기준년월(To)
//			CntcCode = MDC_PS_Common.User_MSTCOD();

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto PH_PY718_Print_Report01_Error;
//			}

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//			WinTitle = "[PH_PY718] 생산완료금액대비O/T현황";

//			ReportName = "PH_PY718_01.rpt";
//			MDC_Globals.gRpt_Formula = new string[2];
//			MDC_Globals.gRpt_Formula_Value = new string[2];
//			MDC_Globals.gRpt_SRptSqry = new string[2];
//			MDC_Globals.gRpt_SRptName = new string[2];
//			MDC_Globals.gRpt_SFormula = new string[2, 2];
//			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//			//// Formula 수식필드

//			//// SubReport


//			MDC_Globals.gRpt_SFormula[1, 1] = "";
//			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

//			/// Procedure 실행"
//			sQry = "                EXEC [PH_PY718_02] ";
//			sQry = sQry + "'" + CLTCOD + "',";
//			//사업장
//			sQry = sQry + "'" + FrYM + "',";
//			//기준년월(Fr)
//			sQry = sQry + "'" + ToYM + "',";
//			//기준년월(To)
//			sQry = sQry + "'" + CntcCode + "'";
//			//조회자 사번

//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", "", 2) == false) {
//				goto PH_PY718_Print_Report01_Error;
//			}

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			return;
//			PH_PY718_Print_Report01_Error:


//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			MDC_Com.MDC_GF_Message(ref "Print_Query_Error:" + Err().Number + " - " + Err().Description, ref "E");

//		}
//	}
//}
