using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 스크랩입고등록
    /// </summary>
    internal class PS_PP048 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP048H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP048L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocEntry;
        private string oStatus;
        private string oCanceled;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP048.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP048_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP048");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1; 
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP048_CreateItems();
                PS_PP048_ComboBox_Setting();
                PS_PP048_CF_ChooseFromList();
                PS_PP048_EnableMenus();
                PS_PP048_SetDocument(oFormDocEntry);
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
        private void PS_PP048_CreateItems()
        {
            try
            {
                oDS_PS_PP048H = oForm.DataSources.DBDataSources.Item("@PS_PP048H");
                oDS_PS_PP048L = oForm.DataSources.DBDataSources.Item("@PS_PP048L");
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
        private void PS_PP048_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                oForm.Items.Item("Div").Specific.ValidValues.Add("10", "스크랩등록");
                oForm.Items.Item("Div").Specific.ValidValues.Add("20", "LOSS등록");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_PP048_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = Convert.ToString(64); //Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_Warehouses);
                oCFLCreationParams.UniqueID = "CFLWAREHOUSES";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oColumn.ChooseFromListUID = "CFLWAREHOUSES";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP048_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP048_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP048_FormItemEnabled();
                    PS_PP048_AddMatrixRow(0, true);
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP048_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP048_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP048_FormClear();
                    oForm.EnableMenu("1281", true);  //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("Div").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Div").Enabled = true;
                    oMat01.Columns.Item("Loss").Editable = true;
                    oMat01.Columns.Item("PP030No").Editable = true;
                    oMat01.Columns.Item("SItemCod").Editable = true;
                    oMat01.Columns.Item("CpCode").Editable = true;
                    oMat01.Columns.Item("ScrapWgt").Editable = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                    oMat01.Columns.Item("Check").Editable = false;
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false);                   //찾기
                    oForm.EnableMenu("1282", true);                    //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("PP030No").Editable = false;
                    oMat01.Columns.Item("ScrapWgt").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                    oMat01.Columns.Item("Check").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true);                    //찾기
                    oForm.EnableMenu("1282", true);                    //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("Div").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat01.Columns.Item("PP030No").Editable = false;
                    oMat01.Columns.Item("SItemCod").Editable = false;
                    oMat01.Columns.Item("CpCode").Editable = false;
                    oMat01.Columns.Item("Loss").Editable = false;
                    oMat01.Columns.Item("ScrapWgt").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                    if (oDS_PS_PP048H.GetValue("CanCeled", 0).ToString().Trim() == "Y")
                    {
                        oMat01.Columns.Item("Check").Editable = false;
                    }
                    else
                    {
                        oMat01.Columns.Item("Check").Editable = true;
                    }
                }
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
        /// 
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP048_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_PP048L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP048L.Offset = oRow;
                oDS_PS_PP048L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
                oForm.Freeze(false);
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
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP048_Validate(string ValidateType)
        {
            bool functionReturnValue = true;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (ValidateType == "검사01")
                {
                }
                else if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        errMessage = "현재모드는 행삭제가 불가능합니다.";
                        throw new Exception();
                    }
                }
                else if (ValidateType == "취소")
                {
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP048_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP048'", "");
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
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP048_DataValidCheck()
        {
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            bool functionReturnValue = true;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP048_FormClear();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "작성일은 필수입니다.";
                    type = "F";
                    ClickCode = "DocDate";
                    throw new Exception();
                }
                if (dataHelpClass.Future_Date_Check(oForm.Items.Item("DocDate").Specific.Value) == "N")
                {
                    errMessage = "미래일자는 입력할 수 없습니다.";
                    type = "F";
                    ClickCode = "DocDate";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("Div").Specific.Value))
                {
                    errMessage = "구분은 필수입니다.";
                    type = "F";
                    ClickCode = "DocDate";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                {
                    errMessage = "담당자는 필수입니다.";
                    type = "F";
                    ClickCode = "CntcCode";
                    throw new Exception();
                }
                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작업지시문서는 필수입니다.";
                        type = "M";
                        ClickCode = "PP030No";
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고코드는 필수입니다.";
                        type = "M";
                        ClickCode = "WhsCode";
                        throw new Exception();
                    }
                    if (oForm.Items.Item("Div").Specific.Value == "10")
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("SItemCod").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "품목코드는 필수입니다.";
                            type = "M";
                            ClickCode = "SItemCod";
                            throw new Exception();
                        }
                    }
                    if (oForm.Items.Item("Div").Specific.Value == "10" && Convert.ToInt32(oMat01.Columns.Item("ScrapWgt").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "스크랩중량은 필수입니다.";
                        type = "M";
                        ClickCode = "ScrapWgt";
                        throw new Exception();
                    }
                    if (oForm.Items.Item("Div").Specific.Value == "10" && Convert.ToInt32(oMat01.Columns.Item("Loss").Cells.Item(i).Specific.Value) > 0)
                    {
                        errMessage = "구분을 스크랩으로 LOSS중량을 입력할 수 없습니다.";
                        type = "M";
                        ClickCode = "Loss";
                        throw new Exception();
                    }
                    if (oForm.Items.Item("Div").Specific.Value == "20" && Convert.ToInt32(oMat01.Columns.Item("Loss").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "LOSS중량은 필수입니다.";
                        type = "M";
                        ClickCode = "Loss";
                        throw new Exception();
                    }
                    if (oForm.Items.Item("Div").Specific.Value == "20" && Convert.ToInt32(oMat01.Columns.Item("ScrapWgt").Cells.Item(i).Specific.Value) > 0)
                    {
                        errMessage = "구분을 LOSS로 스크랩중량을 입력 할 수 없습니다.";
                        type = "M";
                        ClickCode = "ScrapWgt";
                        throw new Exception();
                    }
                }
                if (PS_PP048_Validate("검사01") == false)
                {
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                oDS_PS_PP048L.RemoveRecord(oDS_PS_PP048L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP048_FormClear();
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = false;

                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (type == "M")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP048_Add_InventoryGenEntry
        /// </summary>
        /// <returns></returns>
        private bool PS_PP048_Add_InventoryGenEntry()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int errDiCode;
            int ResultDocNum;
            string errMessage = string.Empty;
            string errDiMsg = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.Comments = "스크랩등록 (" + oDS_PS_PP048H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고_PS_PP048";
                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(j);
                    oDIObject.Lines.ItemCode = oMat01.Columns.Item("SItemCod").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.Quantity = float.Parse(oMat01.Columns.Item("ScrapWgt").Cells.Item(i).Specific.Value);
                }
                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errMessage = "DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg;
                    throw new Exception();
                }
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    oForm.Items.Item("OIGNNo").Specific.Value = ResultDocNum;
                    oDS_PS_PP048H.SetValue("U_OIGNNo", 0, Convert.ToString(ResultDocNum));
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                returnValue = false;
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_PP048_Add_InventoryGenExit
        /// </summary>
        /// <returns></returns>
        private bool PS_PP048_Add_InventoryGenExit()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int ResultDocNum;
            int errDiCode;
            string errDiMsg;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.Comments = "스크랩등록 취소 (" + oDS_PS_PP048H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고_PS_PP048";
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim();

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OIGENum").Cells.Item(i).Specific.Value.ToString().Trim()))
                    {
                        if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
                        {
                            oDIObject.Lines.Add();
                            oDIObject.Lines.SetCurrentLine(j);
                            oDIObject.Lines.ItemCode = oMat01.Columns.Item("SItemCod").Cells.Item(i).Specific.Value;
                            oDIObject.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                            oDIObject.Lines.Quantity = float.Parse(oMat01.Columns.Item("ScrapWgt").Cells.Item(i).Specific.Value);
                        }
                    }
                }
                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errMessage = "DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg;
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        dataHelpClass.DoQuery("UPDATE [@PS_PP048L] SET U_OIGENum = '" + ResultDocNum + "', U_IGE1Num = '" + i + "', U_Check = 'Y' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "' And LineId = '" + oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value + "'");
                        dataHelpClass.DoQuery("UPDATE [@PS_PP040L] SET U_ScrapWt = 0 WHERE DocEntry = '" + oForm.Items.Item("PP040No").Specific.Value + "' And LineId = '" + oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value + "'");//부품 실적추가분 취소처리 => 수량을 0으로 처리
                    }
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                returnValue = false;
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_PP048_MTX01
        /// </summary>
        private void PS_PP048_MTX01()
        {
            int i;
            string Query01;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false); 
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "SELECT 10";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                ProgressBar01.Text = "조회시작!";
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP048L.InsertRecord(i);
                    }
                    oDS_PS_PP048L.Offset = i;
                    oDS_PS_PP048L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP048L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
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
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Link01")
                    {
                        PS_PP040 oTempClass = new PS_PP040();
                        oTempClass.LoadForm(oForm.Items.Item("PP040No").Specific.Value);
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP048_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Items.Item("Div").Specific.Value.ToString().Trim() == "10")
                            {
                                if (PS_PP048_Add_InventoryGenEntry() == false)
                                {
                                    PS_PP048_AddMatrixRow(oMat01.VisualRowCount, false);
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                }
                            }
                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP048_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
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
                                dataHelpClass.DoQuery("EXEC PS_PP048_03 '" + oDocEntry + "'");
                                PS_PP048_FormItemEnabled();
                                PS_PP048_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP048_FormItemEnabled();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", ""); //사용자값활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //사용자값활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP030No"); //사용자값활성(작지문서번호)
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CpCode"); //사용자값활성(공정코드)
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SItemCod"); //사용자값활성(스크랩품목코드)
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "WhsCode"); //사용자값활성(창고코드)

                    if (pVal.ItemUID == "PP030No")
                    {
                        if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
                        {
                            dataHelpClass.MDC_GF_Message("사업장은 필수입니다.", "W");
                            BubbleEvent = false;
                            return;
                        }
                        if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                        {
                            dataHelpClass.MDC_GF_Message("품목코드는 필수입니다.", "W");
                            BubbleEvent = false;
                            return;
                        }
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemChanged == true)
                    {
                        oForm.Freeze(true);
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP048L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP048_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP048H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP048_AddMatrixRow(0, true);
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP048_AddMatrixRow(0, true);

                                oMat01.Columns.Item("BWeight").Visible = true;
                                oMat01.Columns.Item("PWeight").Visible = true;
                                oMat01.Columns.Item("YWeight").Visible = true;
                                oMat01.Columns.Item("NWeight").Visible = true;
                            }
                            else
                            {
                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else
                        {
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "WhsCode")
                        {
                            if (oDataTable01 == null)
                            {
                            }
                            else
                            {
                                oDS_PS_PP048L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_PP048L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
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
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
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
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Check = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" & pVal.Row == 0 & pVal.ColUID == "Check")
                    {
                        oForm.Freeze(true);
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_PP048L.GetValue("U_Check", 0).ToString().Trim()) || oDS_PS_PP048L.GetValue("U_Check", 0).ToString().Trim() == "N")
                        {
                            Check = "Y";
                        }
                        else if (oDS_PS_PP048L.GetValue("U_Check", 0).ToString().Trim() == "Y")
                        {
                            Check = "N";
                        }
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP048L.GetValue("U_OIGENum", 0).ToString().Trim()))
                            {
                                oDS_PS_PP048L.SetValue("U_Check", i, "Y");
                            }
                            else
                            {
                                oDS_PS_PP048L.SetValue("U_Check", i, Check);
                            }
                        }
                        oMat01.LoadFromDataSource();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Query01;
            string YM;
            string MaxDate;
            string DocEntry;
            string ItemCode;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PP030No")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    oDS_PS_PP048L.SetValue("U_PP030No", pVal.Row - 1, "");
                                    oDS_PS_PP048L.SetValue("U_OrdNum", pVal.Row - 1, "");
                                    oDS_PS_PP048L.SetValue("U_PP030HNo", pVal.Row - 1, "");
                                    oDS_PS_PP048L.SetValue("U_PP030MNo", pVal.Row - 1, "");
                                    oDS_PS_PP048L.SetValue("U_CpCode", pVal.Row - 1, "");
                                    oDS_PS_PP048L.SetValue("U_CpName", pVal.Row - 1, "");
                                }
                                else
                                {
                                    Query01 = "Exec PS_PP048_01 '" + oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value + "'";
                                    oRecordSet01.DoQuery(Query01);

                                    if (oRecordSet01.RecordCount == 0)
                                    {
                                        oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                    }
                                    else
                                    {
                                        oDS_PS_PP048L.SetValue("U_PP030No", pVal.Row - 1, oRecordSet01.Fields.Item("PP030No").Value);
                                        oDS_PS_PP048L.SetValue("U_OrdNum", pVal.Row - 1, oRecordSet01.Fields.Item("OrdNum").Value);
                                        oDS_PS_PP048L.SetValue("U_PP030HNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030HNo").Value);
                                        oDS_PS_PP048L.SetValue("U_PP030MNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030MNo").Value);
                                        oDS_PS_PP048L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet01.Fields.Item("CpCode").Value);
                                        oDS_PS_PP048L.SetValue("U_CpName", pVal.Row - 1, oRecordSet01.Fields.Item("CpName").Value);
                                        oDS_PS_PP048L.SetValue("U_WhsCode", pVal.Row - 1, oRecordSet01.Fields.Item("WhsCode").Value);
                                        oDS_PS_PP048L.SetValue("U_WhsName", pVal.Row - 1, oRecordSet01.Fields.Item("WhsName").Value);
                                        oMat01.LoadFromDataSource();
                                    }
                                }
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP048L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP048_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "SItemCod")
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP048L.SetValue("U_SItemNam", pVal.Row - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("SItemCod").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));

                                YM = codeHelpClass.Left(oForm.Items.Item("DocDate").Specific.Value, 6);

                                ItemCode = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value;
                                MaxDate = dataHelpClass.GetValue("SELECT Max(Convert(char(8),U_DocDate,112)) FROM [@PS_MM001H] WHERE Convert(Char(6),U_DocDate,112) = '" + YM + "'", 0, 1);
                                DocEntry = dataHelpClass.GetValue("SELECT Max(DocEntry)  FROM [@PS_MM001H] WHERE U_DocDate = '" + MaxDate + "'", 0, 1);
                                oDS_PS_PP048L.SetValue("U_Price", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_Price FROM [@PS_MM001L] WHERE DocEntry = '" + DocEntry + "' and U_CItemCod = '" + ItemCode + "'", 0, 1));

                                oMat01.LoadFromDataSource();
                            }

                            else if (pVal.ColUID == "ScrapWgt")
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP048L.SetValue("U_ScrapAmt", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(pVal.Row).Specific.Value));
                            }
                            else if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP048L.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat01.Columns.Item("FailCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP048L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP048H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                for (i = 0; i <= oMat01.VisualRowCount; i++)
                                {
                                    oDS_PS_PP048L.SetValue("U_PP030No", i, "");
                                    oDS_PS_PP048L.SetValue("U_OrdNum", i, "");
                                    oDS_PS_PP048L.SetValue("U_PP030HNo", i, "");
                                    oDS_PS_PP048L.SetValue("U_PP030MNo", i, "");
                                    oDS_PS_PP048L.SetValue("U_CpCode", i, "");
                                    oDS_PS_PP048L.SetValue("U_CpName", i, "");
                                    oDS_PS_PP048L.SetValue("U_WhsCode", i, "");
                                    oDS_PS_PP048L.SetValue("U_WhsName", i, "");
                                }
                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);

                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim()))
                                {
                                    oDS_PS_PP048H.SetValue("U_ItemName", 0, "");
                                }
                                else
                                {
                                    oDS_PS_PP048H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'",""));
                                }
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {

                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP048H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP048H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
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
                oForm.Freeze(false);
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
                    PS_PP048_FormItemEnabled();
                    PS_PP048_AddMatrixRow(oMat01.VisualRowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP048H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP048L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    PS_PP048_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        {
            int i;
            double SumQty = 0;
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_PP048_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_PP048L.RemoveRecord(oDS_PS_PP048L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP048_AddMatrixRow(0,false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP048L.GetValue("U_PP030No", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP048_AddMatrixRow(oMat01.RowCount, false);
                            }
                            //합격수량 sum
                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                SumQty += oMat01.Columns.Item("YQty").Cells.Item(i + 1).Specific.Value;
                            }
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int RowCounter = 0;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == true & string.IsNullOrEmpty(oMat01.Columns.Item("OIGENum").Cells.Item(i).Specific.Value.ToString().Trim()))
                                {
                                    RowCounter = RowCounter + 1;
                                }
                            }
                            if (RowCounter == 0)
                            {
                                dataHelpClass.MDC_GF_Message( "취소할 항목을 선택해주세요.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_PP048_Validate("취소") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1"))
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Items.Item("Div").Specific.Value.ToString().Trim() == "10")
                            {
                                if (PS_PP048_Add_InventoryGenExit() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.DoQuery("UPDATE [@PS_PP048L] SET U_Check = 'Y' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'");
                                dataHelpClass.DoQuery("UPDATE [@PS_PP040L] SET U_Loss = 0 WHERE DocEntry = '" + oForm.Items.Item("PP040No").Specific.Value + "'");
                            }
                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE( FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            sQry = "Select Min(IsNULL(U_OIGENum, '')) From [@PS_PP048L] where DocEntry = '" + oDocEntry + "'";
                            oRecordSet01.DoQuery(sQry);
                            if (oForm.Items.Item("Div").Specific.Value.ToString().Trim() == "10")
                            {
                                if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                                {
                                    oStatus = "O";
                                    oCanceled = "N";
                                }
                                else
                                {
                                    oStatus = "C";
                                    oCanceled = "Y";
                                }
                            }
                            else
                            {
                                oStatus = "C";
                                oCanceled = "Y";
                            }

                            dataHelpClass.DoQuery("UPDATE [@PS_PP048H] SET Status = '" + oStatus + "', Canceled = '" + oCanceled + "' WHERE DocEntry = '" + oDocEntry + "'");

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_PP048_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = oDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP048_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP048_FormItemEnabled();
                            PS_PP048_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_PP048_FormItemEnabled();
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
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
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
        }
    }
}
