//using System;

//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    /// <summary>
//    /// 통합재무제표용 계정 관리
//    /// </summary>
//    internal class PS_PP048 : PSH_BaseClass
//    {
//        public string oFormUniqueID;
//        //public SAPbouiCOM.Form oForm;
//        public SAPbouiCOM.Matrix oMat01;
//        private SAPbouiCOM.DBDataSource oDS_PS_PP048H; //등록헤더
//        private SAPbouiCOM.DBDataSource oDS_PS_PP048L; //등록라인

//        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
//        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

//        private string oDocEntry01;

//        private SAPbouiCOM.BoFormMode oFormMode01;

//        /// <summary>
//        /// Form 호출
//        /// </summary>
//        /// <param name="oFromDocEntry01"></param>
//        public override void LoadForm(string oFromDocEntry01)
//        {
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP048.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PS_PP048_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP048");

//                string strXml = null;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                oForm.DataBrowser.BrowseBy = "DocEntry";

//                oForm01.Freeze(true);
//                PS_PP048_CreateItems();
//                PS_PP048_ComboBox_Setting();
//                PS_PP048_CF_ChooseFromList();
//                PS_PP048_EnableMenus();
//                PS_PP048_SetDocument(oFromDocEntry01);
//                PS_PP048_FormResize();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
//            }
//        }

//        /// <summary>
//        /// 화면 Item 생성
//        /// </summary>
//        private void PS_PP048_CreateItems()
//        {
//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// Combobox 설정
//        /// </summary>
//        private void PS_PP048_ComboBox_Setting()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// ChooseFromList
//        /// </summary>
//        private void PS_PP048_CF_ChooseFromList()
//        {
//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// EnableMenus
//        /// </summary>
//        private void PS_PP048_EnableMenus()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// SetDocument
//        /// </summary>
//        /// <param name="oFromDocEntry01">DocEntry</param>
//        private void PS_PP048_SetDocument(string oFromDocEntry01)
//        {
//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// FormResize
//        /// </summary>
//        private void PS_PP048_FormResize()
//        {
//            try
//            {
//                oMat01.AutoResizeColumns();
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// 모드에 따른 아이템 설정
//        /// </summary>
//        private void PS_PP048_FormItemEnabled()
//        {
//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="oRow">행 번호</param>
//        /// <param name="RowIserted">행 추가 여부</param>
//        private void PS_PP048_AddMatrixRow(int oRow, bool RowIserted)
//        {
//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// DocEntry 초기화
//        /// </summary>
//        private void PS_PP048_FormClear()
//        {
//            string DocEntry;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// 필수 사항 check
//        /// </summary>
//        /// <returns></returns>
//        private bool PS_PP048_DataValidCheck()
//        {

//            try
//            {

//            }
//            catch (Exception ex)
//            {
//                if (errCode == "1")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("구분코드가 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                }
//                else if (errCode == "2")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("구분명이 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                }
//                else if (errCode == "3")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                }
//                else if (errCode == "4")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("계정코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oMat01.Columns.Item("AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//                else if (errCode == "5")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("목차제목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oMat01.Columns.Item("Contents").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//            }
//            finally
//            {

//            }

//            return functionReturnValue;
//        }

//        /// <summary>
//        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
//        /// </summary>
//        /// <param name="oUID"></param>
//        /// <param name="oRow"></param>
//        /// <param name="oCol"></param>
//        private void PS_PP048_FlushToItemValue(string oUID, int oRow, string oCol)
//        {
//            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
//            }
//        }

//        /// <summary>
//        /// Form Item Event
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">pVal</param>
//        /// <param name="BubbleEvent">Bubble Event</param>
//        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
//                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
//                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
//                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
//                    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
//                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
//                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
//                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
//                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
//                    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
//                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
//                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
//                    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
//                    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
//                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
//                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
//                    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
//                    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
//                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
//                    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
//                    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
//                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
//                    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
//                    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_Drag: //39
//                    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
//                    break;
//            }
//        }

//        /// <summary>
//        /// ITEM_PRESSED 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// KEY_DOWN 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// GOT_FOCUS 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// COMBO_SELECT 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                oForm.Freeze(true);
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// CLICK 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// VALIDATE 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                BubbleEvent = false;
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// MATRIX_LOAD 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// FORM_UNLOAD 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                    SubMain.Remove_Forms(oFormUniqueID);

//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
//                }
//                else if (pVal.Before_Action == false)
//                {

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// RESIZE 이벤트
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.Before_Action == true)
//                {
//                }
//                else if (pVal.Before_Action == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// EVENT_ROW_DELETE
//        /// </summary>
//        /// <param name="FormUID">Form UID</param>
//        /// <param name="pVal">ItemEvent 객체</param>
//        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
//        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
//        {
//            int i = 0;

//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// FormMenuEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="pVal"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                oForm.Freeze(true);

//                if (pVal.BeforeAction == true)
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284": //취소
//                            break;
//                        case "1286": //닫기
//                            break;
//                        case "1293": //행삭제
//                            break;
//                        case "1281": //찾기
//                            break;
//                        case "1282": //추가
//                            break;
//                        case "1288": //레코드이동(최초)
//                        case "1289": //레코드이동(이전)
//                        case "1290": //레코드이동(다음)
//                        case "1291": //레코드이동(최종)
//                            break;
//                    }
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1284": //취소
//                            break;
//                        case "1286": //닫기
//                            break;
//                        case "1293": //행삭제
//                            break;
//                        case "1281": //찾기
//                        case "1282": //추가
//                            break;
//                        case "1288": //레코드이동(최초)
//                        case "1289": //레코드이동(이전)
//                        case "1290": //레코드이동(다음)
//                        case "1291": //레코드이동(최종)
//                        case "1287": //복제
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        /// <summary>
//        /// FormDataEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="BusinessObjectInfo"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (BusinessObjectInfo.BeforeAction == true)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
//                            break;
//                    }
//                }
//                else if (BusinessObjectInfo.BeforeAction == false)
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        /// <summary>
//        /// RightClickEvent
//        /// </summary>
//        /// <param name="FormUID"></param>
//        /// <param name="pVal"></param>
//        /// <param name="BubbleEvent"></param>
//        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                }

//                switch (pVal.ItemUID)
//                {
//                    case "Mat01":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID01 = pVal.ItemUID;
//                            oLastColUID01 = pVal.ColUID;
//                            oLastColRow01 = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID01 = pVal.ItemUID;
//                        oLastColUID01 = "";
//                        oLastColRow01 = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }
//    }
//}
