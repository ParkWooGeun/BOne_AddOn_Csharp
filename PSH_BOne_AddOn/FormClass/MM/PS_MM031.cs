using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 품의취소등록
    /// </summary>
    internal class PS_MM031 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM031H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM031L; //등록라인
        private string oLastItemUID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM031.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM031_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM031");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM031_CreateItems();
                PS_MM031_SetComboBox();
                PS_MM031_EnableMenus();
                PS_MM031_InitializeForm();
                PS_MM031_SetDocument(oFormDocEntry);
                PS_MM031_FormResize();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PS_MM031_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PS_MM031H = oForm.DataSources.DBDataSources.Item("@PS_MM031H");
                oDS_PS_MM031L = oForm.DataSources.DBDataSources.Item("@PS_MM031L");

                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.Items.Item("BPLID").DisplayDesc = true; //사업장
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
        /// Combobox 설정
        /// </summary>
        private void PS_MM031_SetComboBox()
        {
            try
            {
                //승인여부
                oForm.Items.Item("AdmsYN").Specific.ValidValues.Add("Y", "승인");
                oForm.Items.Item("AdmsYN").Specific.ValidValues.Add("N", "미승인");
                oDS_PS_MM031H.SetValue("U_AdmsYN", 0, "N");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_MM031_InitializeForm
        /// </summary>
        private void PS_MM031_InitializeForm()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                //oDS_PS_MM031H.SetValue("U_BPLID", 0, dataHelpClass.User_BPLID());//아이디별 사업장 세팅
                oDS_PS_MM031H.SetValue("U_RegDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PS_MM031_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); //제거
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1285", true); //복원
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1287", true); //복제
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_MM031_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM031_FormItemEnabled();
                    PS_MM031_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_MM031_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM031_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM031'", "");

                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PS_MM031_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oDS_PS_MM031H.SetValue("U_RegDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅
                    oDS_PS_MM031H.SetValue("U_AdmsYN", 0, "N");

                    oForm.Items.Item("BPLID").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    PS_MM031_FormClear();

                    dataHelpClass.CLTCOD_Select(oForm, "BPLID", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false);  //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BPLID").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "BPLID", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("BPLID").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "BPLID", false); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// 매트릭스 행 추가
        /// </summary>
        private void PS_MM031_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PS_MM031L.GetValue("U_MM030No", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PS_MM031L.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PS_MM031L.InsertRecord(oRow);
                        }
                        oDS_PS_MM031L.Offset = oRow;
                        oDS_PS_MM031L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PS_MM031L.SetValue("U_MM030No", oRow, "");
                        oDS_PS_MM031L.SetValue("U_ItemCode", oRow, "");
                        oDS_PS_MM031L.SetValue("U_ItemName", oRow, "");
                        oDS_PS_MM031L.SetValue("U_ItemSpec", oRow, "");
                        oDS_PS_MM031L.SetValue("U_ItemUnit", oRow, "");
                        oDS_PS_MM031L.SetValue("U_Weight", oRow, "0");
                        oDS_PS_MM031L.SetValue("U_Price", oRow, "0");
                        oDS_PS_MM031L.SetValue("U_Amount", oRow, "0");
                        oDS_PS_MM031L.SetValue("U_Comment", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PS_MM031L.Offset = oRow - 1;
                        oDS_PS_MM031L.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PS_MM031L.SetValue("U_MM030No", oRow - 1, "");
                        oDS_PS_MM031L.SetValue("U_ItemCode", oRow - 1, "");
                        oDS_PS_MM031L.SetValue("U_ItemName", oRow - 1, "");
                        oDS_PS_MM031L.SetValue("U_ItemSpec", oRow - 1, "");
                        oDS_PS_MM031L.SetValue("U_ItemUnit", oRow - 1, "");
                        oDS_PS_MM031L.SetValue("U_Weight", oRow - 1, "0");
                        oDS_PS_MM031L.SetValue("U_Price", oRow - 1, "0");
                        oDS_PS_MM031L.SetValue("U_Amount", oRow - 1, "0");
                        oDS_PS_MM031L.SetValue("U_Comment", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PS_MM031L.Offset = oRow;
                    oDS_PS_MM031L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PS_MM031L.SetValue("U_MM030No", oRow, "");
                    oDS_PS_MM031L.SetValue("U_ItemCode", oRow, "");
                    oDS_PS_MM031L.SetValue("U_ItemName", oRow, "");
                    oDS_PS_MM031L.SetValue("U_ItemSpec", oRow, "");
                    oDS_PS_MM031L.SetValue("U_ItemUnit", oRow, "");
                    oDS_PS_MM031L.SetValue("U_Weight", oRow, "0");
                    oDS_PS_MM031L.SetValue("U_Price", oRow, "0");
                    oDS_PS_MM031L.SetValue("U_Amount", oRow, "0");
                    oDS_PS_MM031L.SetValue("U_Comment", oRow, "");
                    oMat01.LoadFromDataSource();
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
        /// PS_MM031_FormResize
        /// </summary>
        private void PS_MM031_FormResize()
        {
            try
            {
                oForm.Freeze(true);
                oMat01.AutoResizeColumns();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PS_MM031_DataValidCheck()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM031H.GetValue("U_BPLID", 0).ToString().Trim())) //사업장
                {
                    errMessage = "사업장은 필수입니다.";
                    throw new Exception();
                }
                
                if (string.IsNullOrEmpty(oDS_PS_MM031H.GetValue("U_CntcCode", 0).ToString().Trim())) //담당자
                {
                    errMessage = "담당자는 필수입니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PS_MM031H.GetValue("U_CntcCode", 0).ToString().Trim())) //등록일
                {
                    errMessage = "등록일은 필수입니다.";
                    throw new Exception();
                }

                //라인
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인데이터가 없습니다.";
                    throw new Exception();
                }
                
                if (oDS_PS_MM031L.Size > 1)
                {
                    oDS_PS_MM031L.RemoveRecord(oDS_PS_MM031L.Size - 1);
                }

                for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM031L.GetValue("U_Weight", i)) || Convert.ToDouble(oDS_PS_MM031L.GetValue("U_Weight", i).ToString().Trim()) == 0)
                    {
                        errMessage = "수량/중량은 필수입니다.";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM031_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM031_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM031_FormItemEnabled();
                                PS_MM031_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM031_FormItemEnabled();
                                PS_MM031_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM031_FormItemEnabled();
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
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", ""); //담당자
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MM030No"); //품의문서
                }
                else if (pVal.BeforeAction == false)
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                switch (pVal.ItemUID)
                {
                    case "Mat1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
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
                else if (pVal.BeforeAction == false)
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "MM030No") //품의문서
                            {
                                oMat01.FlushToDataSource();
                                //oDS_PS_MM031L.SetValue("U_JobTypNM", pVal.Row - 1, (oDS_PS_MM031L.GetValue("U_JobTyp", pVal.Row - 1).ToString().Trim() == "1" ? "급여" : "상여"));

                                //oDS_PS_MM031L.SetValue("U_JobGbn", pVal.Row - 1, ""); //지급구분
                                //oDS_PS_MM031L.SetValue("U_JobGbnNM", pVal.Row - 1, ""); //지급구분명
                                //oDS_PS_MM031L.SetValue("U_AcctCode", pVal.Row - 1, ""); //관리계정
                                //oDS_PS_MM031L.SetValue("U_AcctName", pVal.Row - 1, ""); //관리계정명
                                //oDS_PS_MM031L.SetValue("U_TeamCode", pVal.Row - 1, ""); //팀코드
                                //oDS_PS_MM031L.SetValue("U_TeamName", pVal.Row - 1, ""); //팀명
                                //oDS_PS_MM031L.SetValue("U_RspCode", pVal.Row - 1, ""); //담당코드
                                //oDS_PS_MM031L.SetValue("U_RspName", pVal.Row - 1, ""); //담당명
                                //oDS_PS_MM031L.SetValue("U_ClsCode", pVal.Row - 1, ""); //반코드
                                //oDS_PS_MM031L.SetValue("U_ClsName", pVal.Row - 1, ""); //반명
                                oMat01.LoadFromDataSource();

                                PS_MM031_AddMatrixRow();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }

                            if (pVal.ColUID == "Weight") //수량
                            {
                                oMat01.FlushToDataSource();
                                //oDS_PS_MM031L.SetValue("U_JobGbnNM", pVal.Row - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oDS_PS_MM031L.GetValue("U_JobGbn", pVal.Row - 1).ToString().Trim() + "'", "AND Code = 'P212'"));
                                oMat01.LoadFromDataSource();

                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }

                            oMat01.AutoResizeColumns();
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            oDS_PS_MM031H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat01.LoadFromDataSource();
                    PS_MM031_FormItemEnabled();
                    PS_MM031_AddMatrixRow();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM031H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM031L);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PS_MM031_FormResize();
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
        /// Raise_EVENT_ROW_DELETE 
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="oMat"></param>
        /// <param name="DBData"></param>
        /// <param name="CheckField"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;
            try
            {
                oForm.Freeze(true);
                if (oLastColRow > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while (i <= DBData.Size - 1)
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord(i);
                                    i = 0;
                                }
                                else
                                {
                                    i += 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
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
            int i;
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
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_MM031_FormItemEnabled();
                            PS_MM031_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PS_MM031_FormItemEnabled();
                            PS_MM031_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            //oDS_PS_MM031H.SetValue("U_RegDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            //oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅
                            //oDS_PS_MM031H.SetValue("U_AdmsYN", 0, "N");
                            PS_MM031_FormItemEnabled();
                            PS_MM031_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PS_MM031_FormItemEnabled();
                            oMat01.AutoResizeColumns();
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat01, oDS_PS_MM031L, "U_LineNum");
                            PS_MM031_AddMatrixRow();
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            oDS_PS_MM031H.SetValue("DocEntry", 0, "");

                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_MM031L.SetValue("DocEntry", i, "");
                                oMat01.LoadFromDataSource();
                            }

                            oForm.Items.Item("BPLID").Enabled = true;
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
    }
}

