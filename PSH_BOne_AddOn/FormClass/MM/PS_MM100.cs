using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 부자재불출등록
    /// </summary>
    internal class PS_MM100 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM100H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM100L; //등록라인

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM100.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM100_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM100");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM100_CreateItems();
                PS_MM100_ComboBox_Setting();
                PS_MM100_CF_ChooseFromList();
                PS_MM100_EnableMenus();
                PS_MM100_SetDocument(oFormDocEntry);
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
        private void PS_MM100_CreateItems()
        {
            try
            {
                oDS_PS_MM100H = oForm.DataSources.DBDataSources.Item("@PS_MM100H");
                oDS_PS_MM100L = oForm.DataSources.DBDataSources.Item("@PS_MM100L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM100_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_MM100_CF_ChooseFromList()
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
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }
            }
        }


        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_MM100_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;

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
                returnValue = true;
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
            }
            return returnValue;
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_MM100_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_MM100_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM100_FormItemEnabled();
                    PS_MM100_AddMatrixRow(0, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM100_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM100_FormClear();
                    oForm.EnableMenu("1281", true);//찾기
                    oForm.EnableMenu("1282", false);//추가
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                }
                if(oForm.Items.Item("OIGENo").Specific.Value != "")
                {
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("Quantity").Editable = false;
                    oMat01.Columns.Item("CCCode").Editable = false;
                }
                else
                {
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("Quantity").Editable = true;
                    oMat01.Columns.Item("CCCode").Editable = true;
                }
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
        /// PS_MM100_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM100_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM100L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM100L.Offset = oRow;
                oDS_PS_MM100L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_MM100_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM100'", "");
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_MM100_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM100_FormClear();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "불출일은 필수입니다.";
                    ClickCode = "DocDate";
                    type = "F";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                {
                    errMessage = "담당자는 필수입니다.";
                    ClickCode = "CntcCode";
                    type = "F";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value, oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목코드는 필수입니다.";
                        ClickCode = "ItemCode";
                        type = "M";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고코드는 필수입니다.";
                        ClickCode = "WhsCode";
                        type = "M";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "불출량은 필수입니다.";
                        ClickCode = "Quantity";
                        type = "M";
                        throw new Exception();
                    }
                    else if (oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value != "103" & oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value != "105")
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("CCCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "불출처는 필수입니다.";
                            ClickCode = "CCCode";
                            type = "M";
                            throw new Exception();
                        }
                    }
                }
                if (PS_MM100_Validate("검사01") == false)
                {
                    returnValue = false;
                    return returnValue;
                }

                oDS_PS_MM100L.RemoveRecord(oDS_PS_MM100L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM100_FormClear();
                }
                returnValue = true;
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
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM100_Add_InventoryGenExit
        /// </summary>
        /// <returns></returns>
        private bool PS_MM100_Add_InventoryGenExit()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int errDiCode = 0;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
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
                oDIObject.Comments = "부자재 불출 등록(" + oDS_PS_MM100H.GetValue("DocEntry", 0).ToString().Trim() + ") 출고 - PS_MM100 ";

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(j);
                    oDIObject.Lines.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.Quantity = Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value);
                    oDIObject.Lines.CostingCode = oMat01.Columns.Item("CCCode").Cells.Item(i).Specific.Value;
                    j += 1;
                }
                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    oDS_PS_MM100H.SetValue("U_OIGENo", 0, Convert.ToString(Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey())));
                    oMat01.LoadFromDataSource();
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                returnValue = false;
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// PS_MM100_Add_InventoryGenEntry
        /// </summary>
        /// <returns></returns>
        private bool PS_MM100_Add_InventoryGenEntry()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int errDiCode = 0;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
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
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGENo").Specific.Value;
                oDIObject.Comments = "부자재 불출 등록 취소(" + oDS_PS_MM100H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고 - PS_MM100 ";

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(j);
                    oDIObject.Lines.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.Quantity = Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value);
                    oDIObject.Lines.CostingCode = oMat01.Columns.Item("CCCode").Cells.Item(i).Specific.Value;
                    j += 1;
                }
                RetVal = oDIObject.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    oDS_PS_MM100H.SetValue("U_OIGNNo", 0, Convert.ToString(Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey())));

                    //dataHelpClass.DoQuery("UPDATE [@PS_MM100H] SET U_OIGNNo = '" + Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey()) + "' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'");
                    oMat01.LoadFromDataSource();
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                returnValue = false;
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Link01")
                    {
                        PS_PP040 PS_PP040 = new PS_PP040();
                        PS_PP040.LoadForm(oForm.Items.Item("PP040No").Specific.Value);
                        BubbleEvent = false;
                        return;
                    }
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM100_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM100_Add_InventoryGenExit() == false)
                            {
                                PS_MM100_AddMatrixRow(oMat01.VisualRowCount, false);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                            }
                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM100_DataValidCheck() == false)
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
                                PS_MM100_FormItemEnabled();
                                PS_MM100_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM100_FormItemEnabled();
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
                    if (pVal.ItemUID == "CntcCode")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                        else if (pVal.ColUID == "TargetDc")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("TargetDc").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            if (pVal.CharPressed == 9)
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 PS_SM010 = new PS_SM010();
                                    PS_SM010.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "WhsCode");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CCCode");
                }
                else if (pVal.Before_Action == false)
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_MM100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_MM100L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM100_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else
                            {
                                oDS_PS_MM100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_MM100H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_MM100H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_MM100_AddMatrixRow(0, true);
                            }
                            else
                            {
                                oDS_PS_MM100H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
            string WhsCode;
            string errMessage = string.Empty;
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
                            if (pVal.ColUID == "ItemCode")
                            {
                                oDS_PS_MM100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                if (!string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                                {
                                    sQry = "Select U_ItmBsort From OITM where ItemCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    if (oRecordSet01.Fields.Item(0).Value == "401")
                                    {
                                        WhsCode = "10" + oForm.Items.Item("BPLId").Specific.Value;
                                        oDS_PS_MM100L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.GetValue("SELECT ItemName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        oDS_PS_MM100L.SetValue("U_OnHand", pVal.Row - 1, dataHelpClass.GetValue("SELECT Sum(OnHand) FROM [OITW] WHERE WhsCode = '" + WhsCode + "' and ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        oMat01.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value = WhsCode;

                                        if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_MM100L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_MM100_AddMatrixRow(pVal.Row, false);
                                        }
                                    }
                                    else
                                    {
                                        errMessage = "부자재만 불출가능합니다.";
                                        throw new Exception();
                                    }
                                }
                                else
                                {
                                    oDS_PS_MM100L.SetValue("U_ItemName", pVal.Row - 1, "");
                                }
                            }
                            else if (pVal.ColUID == "CCCode")
                            {
                                oDS_PS_MM100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_MM100L.SetValue("U_CCName", pVal.Row - 1, dataHelpClass.GetValue("SELECT PrcName FROM OPRC Where PrcCode = '" + oMat01.Columns.Item("CCCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else if (pVal.ColUID == "TargetDc")
                            {
                                sQry = "exec [PS_MM100_01] '" + oForm.Items.Item("BPLId").Specific.Value + "','" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "','" + oMat01.Columns.Item("TargetDc").Cells.Item(pVal.Row).Specific.Value  + "'";
                                oRecordSet01.DoQuery(sQry);

                                oDS_PS_MM100L.SetValue("U_TargetDc", pVal.Row - 1, oRecordSet01.Fields.Item(0).Value);
                                oDS_PS_MM100L.SetValue("U_CntcCode", pVal.Row - 1, oRecordSet01.Fields.Item(5).Value);
                                oDS_PS_MM100L.SetValue("U_CntcName", pVal.Row - 1, oRecordSet01.Fields.Item(6).Value);
                                oDS_PS_MM100L.SetValue("U_UseDept", pVal.Row - 1, oRecordSet01.Fields.Item(7).Value);
                            }
                            else
                            {
                                oDS_PS_MM100L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_MM100H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_MM100H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_MM100H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_MM100H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();

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
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    PS_MM100_FormItemEnabled();
                    PS_MM100_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM100H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM100L);
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                                oDS_PS_MM100L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_MM100L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
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

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" & pVal.Row == 0)
                    {
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
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
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent  pVal, ref bool BubbleEvent)
        {
            try
            {
                int i = 0;
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_MM100_Validate("행삭제") == false)
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
                        oDS_PS_MM100L.RemoveRecord(oDS_PS_MM100L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_MM100_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_MM100L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_MM100_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            if (PS_MM100_Validate("취소") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1"))
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDS_PS_MM100L.RemoveRecord(oDS_PS_MM100L.Size - 1);
                            oMat01.LoadFromDataSource();
                            if (PS_MM100_Add_InventoryGenEntry() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
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
                            sQry = "Select Min(IsNULL(U_OIGNNo, '')) From [@PS_MM100H] where DocEntry = '" + oDocEntry + "'";
                            oRecordSet01.DoQuery(sQry);

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

                            dataHelpClass.DoQuery("UPDATE [@PS_MM100H] SET Status = '" + oStatus + "', Canceled = '" + oCanceled + "' WHERE DocEntry = '" + oDocEntry + "'");

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_MM100_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = oDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_MM100_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_MM100_FormItemEnabled();
                            PS_MM100_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_MM100_FormItemEnabled();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
