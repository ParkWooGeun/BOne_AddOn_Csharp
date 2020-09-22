using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 원가계산재공현황
    /// </summary>
    internal class PS_CO131 : PSH_BaseClass
    {
        public string oFormUniqueID;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO131.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO131_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO131");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                oForm.Freeze(true);
                PS_CO131_CreateItems();
                PS_CO131_ComboBox_Setting();
                PS_CO131_Initialization();
                //PS_CO131_CF_ChooseFromList();
                //PS_CO131_EnableMenus();
                //PS_CO131_SetDocument(oFromDocEntry01);
                //PS_CO131_FormResize();

                oForm.EnableMenu("1283", true); //삭제
                oForm.EnableMenu("1287", true); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
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
        private void PS_CO131_CreateItems()
        {
            try
            {
                oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM").Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO131_Initialization()
        {
            try
            {
                PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
                SAPbouiCOM.ComboBox oCombo = null;

                ////아이디별 사업장 세팅
                oCombo = oForm.Items.Item("BPLId").Specific;
                oCombo.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
        /// PS_CO131_ComboBox_Setting
        /// </summary>
        public void PS_CO131_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                SAPbouiCOM.ComboBox oCombo = null;
                string sQry = null;
                SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //// 사업장
                oCombo = oForm.Items.Item("BPLId").Specific;
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                ////출력구분
                oCombo = oForm.Items.Item("Gbn01").Specific;
                oCombo.ValidValues.Add("1", "제품별");
                oCombo.ValidValues.Add("2", "공정별");
                oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        ///// <summary>
        ///// ChooseFromList
        ///// </summary>
        //private void PS_CO131_CF_ChooseFromList()
        //{
        //    try
        //    {

        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        ///// <summary>
        ///// EnableMenus
        ///// </summary>
        //private void PS_CO131_EnableMenus()
        //{
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        ///// <summary>
        ///// SetDocument
        ///// </summary>
        ///// <param name="oFromDocEntry01">DocEntry</param>
        //private void PS_CO131_SetDocument(string oFromDocEntry01)
        //{
        //    try
        //    {
        //        if (string.IsNullOrEmpty(oFromDocEntry01))
        //        {
        //            PS_CO131_FormItemEnabled();
        //            PS_CO131_AddMatrixRow(0, true);
        //            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else
        //        {
        //            //oForm.Mode = fm_FIND_MODE;
        //            //PS_CO131_FormItemEnabled();
        //            //oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01;
        //            //oForm.Items("1").Click(ct_Regular);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        ///// <summary>
        ///// FormResize
        ///// </summary>
        //private void PS_CO131_FormResize()
        //{
        //    try
        //    {
        //        oMat01.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        ///// <summary>
        ///// 모드에 따른 아이템 설정
        ///// </summary>
        //private void PS_CO131_FormItemEnabled()
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //        {
        //            oForm.Items.Item("Code").Enabled = true;
        //            oForm.Items.Item("Mat01").Enabled = true;
        //            PS_CO131_FormClear();

        //            oForm.EnableMenu("1281", true); //찾기
        //            oForm.EnableMenu("1282", false); //추가
        //        }
        //        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //        {
        //            oForm.Items.Item("Code").Enabled = true;
        //            oForm.Items.Item("Mat01").Enabled = false;

        //            oForm.EnableMenu("1281", false); //찾기
        //            oForm.EnableMenu("1282", true); //추가
        //        }
        //        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //        {
        //            oForm.Items.Item("Code").Enabled = false;
        //            oForm.Items.Item("Mat01").Enabled = true;

        //            oForm.EnableMenu("1281", true); //찾기
        //            oForm.EnableMenu("1282", true); //추가
        //        }

        //        oMat01.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="oRow">행 번호</param>
        ///// <param name="RowIserted">행 추가 여부</param>
        //private void PS_CO131_AddMatrixRow(int oRow, bool RowIserted)
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (RowIserted == false) //행추가여부
        //        {
        //            oDS_PS_CO131L.InsertRecord(oRow);
        //        }

        //        oMat01.AddRow();
        //        oDS_PS_CO131L.Offset = oRow;
        //        oDS_PS_CO131L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
        //        oMat01.LoadFromDataSource();
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        ///// <summary>
        ///// DocEntry 초기화
        ///// </summary>
        //private void PS_CO131_FormClear()
        //{
        //    string DocEntry;
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO131'", "");
        //        if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
        //        {
        //            oForm.Items.Item("DocEntry").Specific.VALUE = 1;
        //        }
        //        else
        //        {
        //            oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        ///// <summary>
        ///// 필수 사항 check
        ///// </summary>
        ///// <returns></returns>
        //private bool PS_CO131_DataValidCheck()
        //{
        //    bool functionReturnValue = false;
        //    int i = 0;
        //    string errCode = string.Empty;

        //    try
        //    {
        //        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //        {
        //            PS_CO131_FormClear();
        //        }

        //        if (string.IsNullOrEmpty(oForm.Items.Item("Code").Specific.VALUE)) //구분코드 미입력
        //        {
        //            errCode = "1";
        //            throw new Exception();
        //        }

        //        if (string.IsNullOrEmpty(oForm.Items.Item("Name").Specific.VALUE)) //구분명 미입력
        //        {
        //            errCode = "2";
        //            throw new Exception();
        //        }

        //        if (oMat01.VisualRowCount == 1) //라인정보 미입력
        //        {
        //            errCode = "3";
        //            throw new Exception();
        //        }

        //        for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
        //        {
        //            if (string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.VALUE))
        //            {
        //                errCode = "4";
        //                throw new Exception();
        //            }

        //            if (string.IsNullOrEmpty(oMat01.Columns.Item("Contents").Cells.Item(i).Specific.VALUE))
        //            {
        //                errCode = "5";
        //                throw new Exception();
        //            }
        //        }

        //        oMat01.FlushToDataSource();
        //        oDS_PS_CO131L.RemoveRecord(oDS_PS_CO131L.Size - 1);
        //        oMat01.LoadFromDataSource();

        //        functionReturnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (errCode == "1")
        //        {
        //            PSH_Globals.SBO_Application.SetStatusBarMessage("구분코드가 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //        else if (errCode == "2")
        //        {
        //            PSH_Globals.SBO_Application.SetStatusBarMessage("구분명이 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //        else if (errCode == "3")
        //        {
        //            PSH_Globals.SBO_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //        else if (errCode == "4")
        //        {
        //            PSH_Globals.SBO_Application.SetStatusBarMessage("계정코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //            oMat01.Columns.Item("AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (errCode == "5")
        //        {
        //            PSH_Globals.SBO_Application.SetStatusBarMessage("목차제목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //            oMat01.Columns.Item("Contents").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {

        //    }

        //    return functionReturnValue;
        //}

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_CO131_FlushToItemValue(string oUID, int oRow= 0 , string oCol = "")
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                string sQry = string.Empty;

                //--------------------------------------------------------------
                //Header--------------------------------------------------------
                switch (oUID)
                {
                    case "ItmBsort":
                        sQry = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code =  '" +oForm.Items.Item("ItmBsort").Specific.VALUE.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oForm.Items.Item("CodeName").Specific.String = oRecordSet01.Fields.Item("Name").Value.ToString().Trim();
                        break;
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

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {

                        System.Threading.Thread thread = new System.Threading.Thread(PS_CO131_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
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
                    if (pVal.CharPressed == 9)
                    {
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

        ///// <summary>
        ///// GOT_FOCUS 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //            if (pVal.ItemUID == "Mat01")
        //            {
        //                if (pVal.Row > 0)
        //                {
        //                    oLastItemUID01 = pVal.ItemUID;
        //                    oLastColUID01 = pVal.ColUID;
        //                    oLastColRow01 = pVal.Row;
        //                }
        //            }
        //            else
        //            {
        //                oLastItemUID01 = pVal.ItemUID;
        //                oLastColUID01 = "";
        //                oLastColRow01 = 0;
        //            }
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        ///// <summary>
        ///// COMBO_SELECT 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        oForm.Freeze(true);
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        ///// <summary>
        ///// CLICK 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //            if (pVal.ItemUID == "Mat01")
        //            {
        //                if (pVal.Row > 0)
        //                {
        //                    oLastItemUID01 = pVal.ItemUID;
        //                    oLastColUID01 = pVal.ColUID;
        //                    oLastColRow01 = pVal.Row;

        //                    oMat01.SelectRow(pVal.Row, true, false);
        //                }
        //            }
        //            else
        //            {
        //                oLastItemUID01 = pVal.ItemUID;
        //                oLastColUID01 = "";
        //                oLastColRow01 = 0;
        //            }
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItmBsort")
                    {
                        PS_CO131_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
            }
        }

        ///// <summary>
        ///// MATRIX_LOAD 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            PS_CO131_FormItemEnabled();
        //            PS_CO131_AddMatrixRow(oMat01.VisualRowCount, false);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

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

        ///// <summary>
        ///// RESIZE 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            PS_CO131_FormResize();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        ///// <summary>
        ///// EVENT_ROW_DELETE
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        //{
        //    int i = 0;

        //    try
        //    {
        //        if (pVal.BeforeAction == true)
        //        {
        //        }
        //        else if (pVal.BeforeAction == false)
        //        {
        //            for (i = 1; i <= oMat01.VisualRowCount; i++)
        //            {
        //                oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //            }

        //            oMat01.FlushToDataSource();
        //            oDS_PS_CO131L.RemoveRecord(oDS_PS_CO131L.Size - 1);
        //            oMat01.LoadFromDataSource();

        //            if (oMat01.RowCount == 0)
        //            {
        //                PS_CO131_AddMatrixRow(0, false);
        //            }
        //            else
        //            {
        //                if (!string.IsNullOrEmpty(oDS_PS_CO131L.GetValue("U_AcctCode", oMat01.RowCount - 1).ToString().Trim()))
        //                {
        //                    PS_CO131_AddMatrixRow(oMat01.RowCount, false);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

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

                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            break;
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            break;
                        case "1285":
                            //복원
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1285":
                            //복원
                            break;
                        case "1293":
                            //행삭제
                            break;
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
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

        ///// <summary>
        ///// 필수입력사항 체크
        ///// </summary>
        ///// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        //private bool PS_CO131_HeaderSpaceLineDel()
        //{
        //    bool functionReturnValue = false;
        //    short ErrNum = 0;

        //    try
        //    {
        //        if (string.IsNullOrEmpty(oDS_PS_CO130H.GetValue("U_YM", 0).ToString().Trim()))
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }

        //        if (string.IsNullOrEmpty(oDS_PS_CO130H.GetValue("U_BPLId", 0).ToString().Trim()))
        //        {
        //            ErrNum = 2;
        //            throw new Exception();
        //        }
        //        functionReturnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("마감년월은 필수입력사항입니다. 확인하세요.");
        //        }
        //        else if (ErrNum == 2)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입력사항입니다. 확인하세요.");
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        functionReturnValue = false;
        //    }
        //    finally
        //    {
        //    }

        //    return functionReturnValue;
        //}

        ///// <summary>
        ///// RightClickEvent
        ///// </summary>
        ///// <param name="FormUID"></param>
        ///// <param name="pVal"></param>
        ///// <param name="BubbleEvent"></param>
        //public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.BeforeAction == true)
        //        {
        //        }
        //        else if (pVal.BeforeAction == false)
        //        {
        //        }

        //        switch (pVal.ItemUID)
        //        {
        //            case "Mat01":
        //                if (pVal.Row > 0)
        //                {
        //                    oLastItemUID01 = pVal.ItemUID;
        //                    oLastColUID01 = pVal.ColUID;
        //                    oLastColRow01 = pVal.Row;
        //                }
        //                break;
        //            default:
        //                oLastItemUID01 = pVal.ItemUID;
        //                oLastColUID01 = "";
        //                oLastColRow01 = 0;
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}


        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_CO131_Print_Report01()
        {

            string WinTitle = null;
            string ReportName = null;

            string BPLId = string.Empty;
            string YM = string.Empty;
            string ItmBSort = string.Empty;
            string Gbn01 = string.Empty;

            string DocDate = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLId = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.VALUE.ToString().Trim();
                ItmBSort = oForm.Items.Item("ItmBsort").Specific.VALUE.ToString().Trim();
                Gbn01 = oForm.Items.Item("Gbn01").Specific.VALUE.ToString().Trim();

                WinTitle = "원가계산재공현황[PS_CO131]";
                if (Gbn01 == "1")
                {
                    ReportName = "PS_CO131_01.RPT";
                }
                else if (Gbn01 == "2")
                {
                    ReportName = "PS_CO131_02.RPT";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLId, ""))); //사업장
                dataPackFormula.Add(new PSH_DataPackClass("@YM", YM.Substring(0, 4) + "년 " + YM.Substring(4, 2) + "월")); //년월

                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId)); //일자
                dataPackParameter.Add(new PSH_DataPackClass("@YM", YM)); //일자
                dataPackParameter.Add(new PSH_DataPackClass("@ItemBsort", ItmBSort)); //일자

                formHelpClass.CrystalReportOpen(WinTitle, ReportName,dataPackParameter, dataPackFormula);
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
