using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.CodeDom;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기간비용등록
    /// </summary>
    internal class PS_CO660 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;

        private SAPbouiCOM.DBDataSource oDS_PS_CO660H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO660L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocType01;
        private string oDocEntry01;
        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO660.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO660_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO660");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_CO660_CreateItems();
                PS_CO660_ComboBox_Setting();
                PS_CO660_EnableMenus();
                PS_CO660_SetDocument(oFromDocEntry01);
                
                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", true); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
                oForm.EnableMenu("1299", false); //행닫기
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
        private void PS_CO660_CreateItems()
        {
            try
            {
                oDS_PS_CO660H = oForm.DataSources.DBDataSources.Item("@PS_CO660H");
                oDS_PS_CO660L = oForm.DataSources.DBDataSources.Item("@PS_CO660L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO660_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false); //사업장

                //종결여부
                oMat01.Columns.Item("EndYN").ValidValues.Add("N", "N");
                oMat01.Columns.Item("EndYN").ValidValues.Add("Y", "Y");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_CO660_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_CO660_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PS_CO660_FormItemEnabled();
                    PS_CO660_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_CO660_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO660_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    PS_CO660_FormClear();

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가

                    oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = "";
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;

                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
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
        /// 매트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO660_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_CO660L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_CO660L.Offset = oRow;
                oDS_PS_CO660L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                oMat01.LoadFromDataSource();
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_CO660_FormClear()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO660'", "");

                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_CO660_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errCode = string.Empty;

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_CO660_FormClear();
                }

                //사업장 미입력 시
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLID").Specific.Selected.Value))
                {
                    errCode = "1";
                    throw new Exception();
                }

                //계정 미입력 시
                if (string.IsNullOrEmpty(oForm.Items.Item("AcctCode").Specific.VALUE))
                {
                    errCode = "2";
                    throw new Exception();
                }

                //기간(시작) 미입력 시
                if (string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.VALUE))
                {
                    errCode = "3";
                    throw new Exception();
                }

                //기간(종료) 미입력시
                if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.VALUE))
                {
                    errCode = "4";
                    throw new Exception();
                }

                //금액 미입력시
                if (string.IsNullOrEmpty(oForm.Items.Item("Expense").Specific.VALUE))
                {
                    errCode = "5";
                    throw new Exception();
                }

                //라인정보 미입력 시
                if (oMat01.VisualRowCount < 1)
                {
                    errCode = "500";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    //프로그램구분 미입력 시
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("Expense").Cells.Item(i).Specific.VALUE))
                    {
                        errCode = "501";
                        throw new Exception();
                    }
                }

                oMat01.FlushToDataSource();
                oDS_PS_CO660L.RemoveRecord(oDS_PS_CO660L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_CO660_FormClear();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("계정과목을 입력하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기간(시작)을 입력하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기간(종료)를 입력하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("금액을 입력하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "500")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "501")
                {
                    oMat01.Columns.Item("Expense").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    PSH_Globals.SBO_Application.StatusBar.SetText("금액은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 금액과 라인의 합계 비교
        /// </summary>
        /// <returns></returns>
        private bool PS_CO660_ExpCheck()
        {
            bool returnValue = false;
            string errCode = string.Empty;

            try
            {
                if (oForm.Items.Item("Expense").Specific.VALUE != oForm.Items.Item("TotalExp").Specific.VALUE)
                {
                    errCode = "1";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("금액과 라인합계가 일치하지 않습니다. 확인하십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 메트릭스에 데이터 로드
        /// </summary>
        private void PS_CO660_MTX01()
        {
            int i;
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            double TotalExp = 0;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                Param01 = oForm.Items.Item("FrDate").Specific.Value;
                Param02 = oForm.Items.Item("ToDate").Specific.Value;
                Param03 = oForm.Items.Item("Expense").Specific.Value;
                
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                oForm.Freeze(true);

                sQry = "EXEC PS_CO660_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oMat01.Clear();
                    errCode = "1";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_CO660L.InsertRecord(i);
                    }
                    
                    if (i == oRecordSet.RecordCount) //마지막 빈행 추가
                    {
                        oDS_PS_CO660L.Offset = i;
                        oDS_PS_CO660L.SetValue("U_LineNum", i, Convert.ToString(i + 1)); //라인번호
                        oDS_PS_CO660L.SetValue("U_StdYM", i, ""); //기준년월
                        oDS_PS_CO660L.SetValue("U_StdDCnt", i, ""); //기준일수
                        oDS_PS_CO660L.SetValue("U_Expense", i, Convert.ToString(0)); //금액
                        oDS_PS_CO660L.SetValue("U_Comments", i, ""); //비고
                        oDS_PS_CO660L.SetValue("U_EndYN", i, ""); //종결여부
                        oDS_PS_CO660L.SetValue("U_JEYN", i, ""); //분개여부
                    }
                    else
                    {
                        oDS_PS_CO660L.Offset = i;
                        oDS_PS_CO660L.SetValue("U_LineNum", i, Convert.ToString(i + 1)); //라인번호
                        oDS_PS_CO660L.SetValue("U_StdYM", i, oRecordSet.Fields.Item("StdYM").Value); //기준년월
                        oDS_PS_CO660L.SetValue("U_StdDCnt", i, oRecordSet.Fields.Item("StdDCnt").Value); //기준일수
                        oDS_PS_CO660L.SetValue("U_Expense", i, oRecordSet.Fields.Item("Expense").Value); //금액
                        oDS_PS_CO660L.SetValue("U_Comments", i, oRecordSet.Fields.Item("Comments").Value); //비고
                        oDS_PS_CO660L.SetValue("U_EndYN", i, oRecordSet.Fields.Item("EndYN").Value); //종결여부
                        oDS_PS_CO660L.SetValue("U_JEYN", i, oRecordSet.Fields.Item("JEYN").Value); //분개여부

                        oRecordSet.MoveNext();
                    }

                    TotalExp += Convert.ToDouble(oDS_PS_CO660L.GetValue("U_Expense", i));

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

                }

                oDS_PS_CO660H.SetValue("U_TotalExp", 0, Convert.ToString(TotalExp));

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                ProgressBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
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

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                    if (pVal.ItemUID == "PS_CO660")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_CO660_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_CO660_ExpCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_CO660_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "btnCal")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_CO660_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            PS_CO660_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PS_CO660")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_CO660_FormItemEnabled();
                                PS_CO660_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_CO660_FormItemEnabled();
                            }
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "PGCd")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PGCd");
                        }
                    }
                    else if (pVal.ItemUID == "AcctCode") //계정과목
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode", "");
                    }
                    else if (pVal.ItemUID == "CoAcctCD") //상대계정
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CoAcctCD", "");
                    }
                    else if (pVal.ItemUID == "OcrCode") //배부규칙
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OcrCode", "");
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        oMat01.FlushToDataSource();
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
            double TotalExp = 0;
            string Query01 = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "Expense")
                            {
                                for (int loopCount = 1; loopCount <= oMat01.VisualRowCount; loopCount++)
                                {
                                    TotalExp += Convert.ToDouble(oMat01.Columns.Item("Expense").Cells.Item(loopCount).Specific.Value);
                                }

                                oForm.Items.Item("TotalExp").Specific.Value = TotalExp;
                            }
                        }
                        
                        if (pVal.ItemUID == "AcctCode") //계정
                        {
                            Query01 = "SELECT AcctName FROM OACT WHERE AcctCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            RecordSet01.DoQuery(Query01);
                            oForm.Items.Item("AcctName").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "CoAcctCD") //상대계정
                        {
                            Query01 = "SELECT AcctName FROM OACT WHERE AcctCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            RecordSet01.DoQuery(Query01);
                            oForm.Items.Item("CoAcctNM").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "OcrCode") //배부규칙
                        {
                            Query01 = "SELECT OcrName FROM OOCR WHERE OcrCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            RecordSet01.DoQuery(Query01);
                            oForm.Items.Item("OcrName").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        oForm.Update();
                        oMat01.AutoResizeColumns();
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO660_FormItemEnabled();
                    PS_CO660_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
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
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
        
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pval, bool BubbleEvent)
        {
            int i = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pval.BeforeAction == true)
                    {
                    }
                    else if (pval.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }

                        oMat01.FlushToDataSource();
                        oDS_PS_CO660L.RemoveRecord(oDS_PS_CO660L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_CO660_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_CO660L.GetValue("U_PGType", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_CO660_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_CO660_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_CO660_FormItemEnabled();
                            PS_CO660_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_CO660_FormItemEnabled();
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
            finally
            {
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
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
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
