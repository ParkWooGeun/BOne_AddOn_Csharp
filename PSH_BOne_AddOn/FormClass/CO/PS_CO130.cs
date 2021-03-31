﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Timers;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 제품별 원가계산
    /// </summary>
    internal class PS_CO130 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_CO130H;  //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO130L;  //등록라인

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO130.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_CO130_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO130");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.Freeze(true);
                PS_CO130_CreateItems();
                PS_CO130_ComboBox_Setting();

                oForm.EnableMenu("1283", true);               //// 삭제
                oForm.EnableMenu("1287", true);               //// 복제
                oForm.EnableMenu("1286", false);              //// 닫기
                oForm.EnableMenu("1284", false);              //// 취소
                oForm.EnableMenu("1293", true);               //// 행삭제
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PS_CO130_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                //디비데이터 소스 개체 할당
                oDS_PS_CO130H = oForm.DataSources.DBDataSources.Item("@PS_CO130H");
                oDS_PS_CO130L = oForm.DataSources.DBDataSources.Item("@PS_CO130L");

                oMat01 = oForm.Items.Item("Mat01").Specific; //메트릭스 개체 할당
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
        /// 콤보박스 Setting
        /// </summary>
        private void PS_CO130_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// AddOn 연결 유지용 Timer 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void KeepAddOnConnection(object sender, ElapsedEventArgs e)
        {
            PSH_Globals.SBO_Application.RemoveWindowsMessage(BoWindowsMessageType.bo_WM_TIMER, true);
        }

        /// <summary>
        /// 제품별원가계산 결과 저장
        /// </summary>
        private void SaveData()
        {
            string sQry;
            string YM;
            string BPLID;
            string UserSign;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            Timer timer = new Timer();

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                oForm.Freeze(true);

                YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

                sQry = "EXEC [PS_CO130_50] '";
                sQry += YM + "','";
                sQry += BPLID + "','";
                sQry += UserSign + "'";

                timer.Interval = 30000; //30초
                timer.Elapsed += KeepAddOnConnection;
                timer.Start();

                oRecordSet01.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                timer.Stop();
                timer.Dispose();

                oForm.Update();
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 제품별원가계산 결과 저장 후 문서 찾기
        /// </summary>
        private void FindForm()
        {
            string BPLID;
            string StdYM;

            try
            {
                oForm.Freeze(true);

                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                StdYM = oForm.Items.Item("YM").Specific.Value;

                //찾기모드 변경
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                oForm.Items.Item("BPLId").Specific.Select(BPLID);
                oForm.Items.Item("YM").Specific.Value = StdYM;

                oForm.Items.Item("1").Click();
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

        #region SaveData 메소드 : 제품별원가계산 결과 저장(ADO.NET, 외부 프로그램 실행 테스트 참조) 절대 삭제 금지(2020.10.23 송명규)
        ///// <summary>
        ///// 제품별원가계산 결과 저장(ADO.NET, 외부 프로그램 실행 테스트 참조) 절대 삭제 금지(2020.10.23 송명규)
        ///// </summary>
        //private void SaveData()
        //{
        //    string sQry;
        //    string YM;
        //    string BPLID;
        //    string UserSign;
        //    string errCode = string.Empty;

        //    PSH_DatabaseHelpClass databaseHelpClass;
        //    Timer timer = new Timer();

        //    try
        //    {
        //        databaseHelpClass = new PSH_DatabaseHelpClass(PSH_Globals.SBO_Application.Company.ServerName, PSH_Globals.oCompany.CompanyDB, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW);

        //        Process.Start(PSH_Globals.SP_Path + "\\ExtProgram\\WaitingForm.exe"); //"처리중" 외부 ALERT 창 실행

        //        timer.Interval = 30000; //30초
        //        timer.Elapsed += KeepAddOnConnection;
        //        timer.Start();

        //        YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
        //        BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
        //        UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

        //        sQry = "EXEC [PS_CO130_50] '";
        //        sQry += YM + "','";
        //        sQry += BPLID + "','";
        //        sQry += UserSign + "'";

        //        int excuteQueryReturn = databaseHelpClass.ExecuteNonQuery(sQry, System.Data.CommandType.Text); //ADO.NET 사용
        //    }
        //    catch(Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        timer.Stop();
        //        timer.Dispose();

        //        foreach (Process process in Process.GetProcesses()) //"처리중" 외부 ALERT 창 종료
        //        {
        //            if (process.ProcessName.StartsWith("WaitingForm"))
        //            {
        //                process.Kill();
        //            }
        //        }
        //    }
        //}
        #endregion

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PS_CO130_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            short ErrNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim()))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PS_CO130H.GetValue("U_BPLId", 0).ToString().Trim()))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("마감년월은 필수입력사항입니다. 확인하세요.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입력사항입니다. 확인하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        private bool PS_CO130_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            short ErrNum = 0;

            try
            {
                oMat01.FlushToDataSource();

                //라인
                if (oMat01.VisualRowCount == 0)
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                oMat01.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                int i = 0;
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            oDS_PS_CO130H.SetValue("Code", 0, "");
                            oDS_PS_CO130H.SetValue("Name", 0, "");
                            oDS_PS_CO130H.SetValue("U_YM", 0, "");
                            oDS_PS_CO130H.SetValue("U_BPLId", 0, "");

                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_CO130L.SetValue("Code", i, "");
                                oMat01.LoadFromDataSource();
                            }

                            oForm.Freeze(false);
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string BPLID;
            string YM;
            string Code;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_CO130_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_CO130_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            YM = oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim();
                            BPLID = oDS_PS_CO130H.GetValue("U_BPLId", 0).ToString().Trim();
                            Code = YM + BPLID;
                            oDS_PS_CO130H.SetValue("Code", 0, Code);
                            oDS_PS_CO130H.SetValue("Name", 0, Code);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_CO130_HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        //전기기간 체크
                        PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
                        
                        if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim().Substring(0, 4) + "-" + oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim().Substring(4, 2) + "'", "") == "Y") //전기기간 잠김
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("해당월의 전기기간이 잠겼습니다. 확인하십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }

                        //당월 자료 존재 여부 체크
                        if (Convert.ToInt32(dataHelpClass.Get_ReData("COUNT(*)", "Code", "[@PS_CO130H]", "'" + oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim() + oDS_PS_CO130H.GetValue("U_BPLId", 0).ToString().Trim() + "'", "")) > 0)
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("해당월의 자료가 존재합니다. 확인하십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }

                        SaveData();
                        FindForm();
                    }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO130H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO130L);
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

    }
}
