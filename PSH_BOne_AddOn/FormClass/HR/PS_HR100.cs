using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사규위반조치등록
    /// </summary>
    internal class PS_HR100 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_HR100H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_HR100L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR100.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_HR100_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_HR100");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_HR100_CreateItems();
                PS_HR100_ComboBox_Setting();
                PS_HR100_EnableMenus();
                PS_HR100_SetDocument(oFromDocEntry01);
                PS_HR100_FormResize();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1285", false); //복원
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PS_HR100_CreateItems()
        {
            try
            {
                oDS_PS_HR100H = oForm.DataSources.DBDataSources.Item("@PS_HR100H");
                oDS_PS_HR100L = oForm.DataSources.DBDataSources.Item("@PS_HR100L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_HR100_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                sQry = " SELECT A.U_GrpCode,";
                sQry += "       A.U_GrpName";
                sQry += " FROM [@PS_HR000H] AS A";
                sQry += " WHERE A.Canceled = 'N'";
                sQry += " ORDER BY A.U_GrpCode";
                oForm.Items.Item("GrpCode").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("GrpCode").Specific, sQry, "", false, false);
                oForm.Items.Item("GrpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                sQry = "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'H001'";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ActGd"), sQry,"", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_HR100_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_HR100_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_HR100_FormItemEnabled();
                    PS_HR100_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_HR100_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_HR100_FormResize()
        {
            try
            {
                oForm.Freeze(true);
                oMat01.AutoResizeColumns();
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_HR100_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("GrpCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;

                    PS_HR100_FormClear();
                    oForm.EnableMenu("1281", true);
                    oForm.EnableMenu("1282", false);

                    oForm.Items.Item("DocDate").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CrCode").Specific.VALUE = dataHelpClass.User_MSTCOD();

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {

                    oForm.Items.Item("GrpCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Specific.VALUE = "";
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.EnableMenu("1281", false);
                    oForm.EnableMenu("1282", true);

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("GrpCode").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                }
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
        /// PS_HR100_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_HR100_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)//행추가여부
                {
                    oDS_PS_HR100L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_HR100L.Offset = oRow;
                oDS_PS_HR100L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_HR100_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR100'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_HR100_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_HR100_FormClear();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("GrpCode").Specific.VALUE))
                {
                    errMessage = "조치구분이 선택되지 않았습니다.";
                    ClickCode = "GrpCode";
                    type = "F";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CrCode").Specific.VALUE))
                {
                    errMessage = "등록자가 입력되지 않았습니다.";
                    ClickCode = "CrCode";
                    type = "F";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE))
                {
                    errMessage = "등록일이 입력되지 않았습니다.";
                    ClickCode = "DocDate";
                    type = "F";
                    throw new Exception();
                }
                
                if (oMat01.VisualRowCount == 1) //라인정보 미입력 시
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    type = "X";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("VioCode").Cells.Item(i).Specific.VALUE))//위반자사번
                    {
                        errMessage = "위반자는 필수입니다.";
                        ClickCode = "VioCode";
                        type = "M";
                        throw new Exception();
                    }
                    
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("Location").Cells.Item(i).Specific.VALUE))//장소
                    {
                        errMessage = "장소는 필수입니다.";
                        ClickCode = "Location";
                        type = "M";
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("CodeLv1").Cells.Item(i).Specific.VALUE))  //항목
                    {
                        errMessage = "항목은 필수입니다.";
                        ClickCode = "CodeLv1";
                        type = "M";
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("GrpDate").Cells.Item(i).Specific.VALUE))//조치일자
                    {
                        errMessage = "조치일자는 필수입니다.";
                        ClickCode = "GrpDate";
                        type = "M";
                        throw new Exception();
                    }
                }

                oMat01.FlushToDataSource();
                oDS_PS_HR100L.RemoveRecord(oDS_PS_HR100L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_HR100_FormClear();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {

                    if (type == "F")
                    {
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else if (type == "M")
                    {
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_HR100_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();
                            oFormMode01 = oForm.Mode;

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_HR100_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.VALUE.ToString().Trim();
                            oFormMode01 = oForm.Mode;
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
                                PS_HR100_FormItemEnabled();
                                PS_HR100_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_HR100_FormItemEnabled();
                            }
                        }
                    }
                }
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "VioCode")
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "VioCode");
                        }
                        else if (pVal.ColUID == "CodeLv1")
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CodeLv1");
                        }
                        else if (pVal.ColUID == "CodeLv2")
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CodeLv2");
                        }
                        else if (pVal.ColUID == "CodeLv3")
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CodeLv3");
                        }
                    }
                    else if (pVal.ItemUID == "CrCode")
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CrCode", "");
                    }
                }
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
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
                if (pVal.BeforeAction == true)
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
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        oMat01.FlushToDataSource();

                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "VioCode")
                            {
                                oDS_PS_HR100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);

                                sQry = "EXEC PS_HR100_01 '";
                                sQry += oForm.Items.Item("BPLId").Specific.VALUE + "','";
                                sQry += oForm.Items.Item("GrpCode").Specific.VALUE + "','";
                                sQry += oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'";

                                oRecordSet01.DoQuery(sQry);
                                oDS_PS_HR100L.SetValue("U_VioName", pVal.Row - 1, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim());
                                oDS_PS_HR100L.SetValue("U_ActGd", pVal.Row - 1, oRecordSet01.Fields.Item("NextActGd").Value.ToString().Trim());

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_HR100L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_HR100_AddMatrixRow(pVal.Row, false);
                                }

                            }
                            else if (pVal.ColUID == "CodeLv1")
                            {
                                sQry = "SELECT B.U_NameLv1";
                                sQry += " FROM [@PS_HR000H] AS A INNER JOIN [@PS_HR000L] AS B ON A.DocEntry = B.DocEntry";
                                sQry += " WHERE A.Canceled = 'N'";
                                sQry += "   AND B.U_UseYN = 'Y'";
                                sQry += "   AND A.U_GrpCode = '" + oForm.Items.Item("GrpCode").Specific.VALUE + "'";
                                sQry += "   AND B.U_CodeLv1 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'";

                                oRecordSet01.DoQuery(sQry);
                                oDS_PS_HR100L.SetValue("U_NameLv1", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                            }
                            else if (pVal.ColUID == "CodeLv2")
                            {

                                sQry = "SELECT B.U_NameLv2";
                                sQry += " FROM [@PS_HR001H] AS A INNER JOIN [@PS_HR001L] AS B ON A.DocEntry = B.DocEntry";
                                sQry += " WHERE    A.Canceled = 'N'";
                                sQry += "   AND B.U_UseYN = 'Y'";
                                sQry += "   AND A.U_GrpCode = '" + oForm.Items.Item("GrpCode").Specific.VALUE + "'";
                                sQry += "   AND A.U_CodeLv1 = '" + oMat01.Columns.Item("CodeLv1").Cells.Item(pVal.Row).Specific.VALUE + "'";
                                sQry += "   AND B.U_CodeLv2 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'";

                                oRecordSet01.DoQuery(sQry);
                                oDS_PS_HR100L.SetValue("U_NameLv2", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                            }
                            else if (pVal.ColUID == "CodeLv3")
                            {

                                sQry = "SELECT B.U_NameLv3";
                                sQry += " FROM [@PS_HR002H] AS A INNER JOIN [@PS_HR002L] AS B ON A.DocEntry = B.DocEntry";
                                sQry += " WHERE A.Canceled = 'N'";
                                sQry += "   AND B.U_UseYN = 'Y'";
                                sQry += "   AND A.U_GrpCode = '" + oForm.Items.Item("GrpCode").Specific.VALUE + "'";
                                sQry += "   AND A.U_CodeLv1 = '" + oMat01.Columns.Item("CodeLv1").Cells.Item(pVal.Row).Specific.VALUE + "'";
                                sQry += "   AND A.U_CodeLv2 = '" + oMat01.Columns.Item("CodeLv2").Cells.Item(pVal.Row).Specific.VALUE + "'";
                                sQry += "   AND B.U_CodeLv3 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE + "'";

                                oRecordSet01.DoQuery(sQry);
                                oDS_PS_HR100L.SetValue("U_NameLv3", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                            }
                            else
                            {
                                oDS_PS_HR100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oForm.Update();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "CrCode")
                        {
                            sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("CrName").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    PS_HR100_FormItemEnabled();
                    PS_HR100_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    PS_HR100_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_HR100L.RemoveRecord(oDS_PS_HR100L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_HR100_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_HR100L.GetValue("U_CodeLv3", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_HR100_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_HR100_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_HR100_FormItemEnabled();
                            PS_HR100_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_HR100_FormItemEnabled();
                            break;
                        case "1287": //복제
                            break;
                    }
                }
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
