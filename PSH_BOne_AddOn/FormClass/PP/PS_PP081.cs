using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 부품-생산완료등록
    /// </summary>
    internal class PS_PP081 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_PP081H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP081L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP0811L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocEntry;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP081.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP081_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP081");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP081_CreateItems();
                PS_PP081_ComboBox_Setting();
                PS_PP081_CF_ChooseFromList();
                PS_PP081_EnableMenus();
                PS_PP081_SetDocument(oFormDocEntry);
                PS_PP081_FormResize();
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
        private void PS_PP081_CreateItems()
        {
            try
            {
                oDS_PS_PP081H = oForm.DataSources.DBDataSources.Item("@PS_PP080H");
                oDS_PS_PP081L = oForm.DataSources.DBDataSources.Item("@PS_PP080L");
                oDS_PS_PP0811L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP081_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE Code In ('102','602') And U_PudYN = 'Y' order by Code", "", false, false);
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_PP081_CF_ChooseFromList()
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
        /// EnableMenus
        /// </summary>
        private void PS_PP081_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false); //메뉴설정
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP081_MTX01
        /// </summary>
        private void PS_PP081_MTX01()
        {
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();

                Query01 = "SELECT 10";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다";
                    throw new Exception();
                }

                ProgressBar01.Text = "조회시작!";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP081L.InsertRecord(i);
                    }
                    oDS_PS_PP081L.Offset = i;
                    oDS_PS_PP081L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP081L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
            }
        }

        /// <summary>
        /// PS_PP081_Add_InventoryGenExit
        /// </summary>
        /// <returns></returns>
        private bool PS_PP081_Add_InventoryGenEntry()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int errDiCode = 0;
            int ResultDocNum;
            string errCode = string.Empty;
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
                oDIObject.Comments = "부품생산완료등록 (" + oDS_PS_PP081H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고 - PS_PP081 ";

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    oDIObject.Lines.Add();
                    oDIObject.Lines.SetCurrentLine(j);
                    oDIObject.Lines.ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.WarehouseCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
                    oDIObject.Lines.Quantity = float.Parse(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value);
                    //부품,멀티인경우 배치를 선택
                    if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value == "102" || oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value == "104" || oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value == "111")
                    {
                        if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value) == "Y") //배치사용품목이면
                        {
                            oDIObject.Lines.BatchNumbers.BatchNumber = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                            oDIObject.Lines.BatchNumbers.Quantity = float.Parse(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value);
                            oDIObject.Lines.BatchNumbers.Add();
                        }
                    }
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
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    oForm.Items.Item("OIGNNo").Specific.Value = ResultDocNum;
                    oDS_PS_PP081H.SetValue("U_OIGNNo", 0, Convert.ToString(ResultDocNum));
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
        /// PS_PP081_Validate
        /// </summary>
        /// <returns></returns>
        private bool PS_PP081_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            int i;
            string Query01;
            string OIGNNo;
            string OrdGbn;
            string errMessage = string.Empty;
            string errCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP080H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.";
                    throw new Exception();
                }

                if (ValidateType == "검사01")
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107")
                        {
                            if (Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "작업지시문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }
                    }
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
                    OIGNNo = oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim();
                    OrdGbn = oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim();
                    Query01 = "EXEC PS_PP080_04 '" + OrdGbn + "','" + OIGNNo + "'";
                    oRecordSet01.DoQuery(Query01);
                    for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                    {
                        errMessage = "품목코드: " + oRecordSet01.Fields.Item(0).Value + " 의 재고가 부족합니다.";
                        throw new Exception();
                    }
                }
                functionReturnValue = true;
            }            
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP081_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP081_FormItemEnabled();
                    PS_PP081_AddMatrixRow(0, true);
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
        private void PS_PP081_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat01").Top = 82;
                oForm.Items.Item("Mat01").Left = 6;
                oForm.Items.Item("Mat01").Width = oForm.Width - 18;
                oForm.Items.Item("Mat01").Height = (oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height / 2 - 10));

                oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat01").Height + oForm.Items.Item("Mat01").Top + 35;
                oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat01").Left;
                oForm.Items.Item("Mat02").Width = oForm.Items.Item("Mat01").Width;
                oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height - 5;

                oForm.Items.Item("1").Top = oForm.Items.Item("Mat02").Top - 30;
                oForm.Items.Item("2").Top = oForm.Items.Item("Mat02").Top - 30;

                oForm.Items.Item("27").Top = oForm.Items.Item("Mat02").Top - 30;
                oForm.Items.Item("27").Left = oForm.Items.Item("2").Left + 200;

                oForm.Items.Item("SumQty").Top = oForm.Items.Item("Mat02").Top - 30;
                oForm.Items.Item("SumQty").Left = oForm.Items.Item("27").Left + 60;

                oForm.Items.Item("30").Top = oForm.Items.Item("Mat02").Top - 20;
                oForm.Items.Item("30").Left = oForm.Items.Item("Mat02").Width - 150;

                oForm.Items.Item("Focus").Top = oForm.Items.Item("Mat02").Top - 30;
                oForm.Items.Item("Focus").Left = oForm.Items.Item("Mat02").Width - 20;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP081_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP081_FormClear();
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("SumQty").Specific.Value = 0;//수량 Sum
                } 
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("SumQty").Specific.Value = 0;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("OrdGbn").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();
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
        /// PS_PP081_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP081_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP081L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP081L.Offset = oRow;
                oDS_PS_PP081L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_PP081_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP080'", "");
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
        private bool PS_PP081_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string BPLID;
            string OrdGbn;
            string PP030HNo;
            string ORDRNo;
            string RDR1No;
            string Query01;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP081_FormClear();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "작성일은 필수입니다.";
                    ClickCode = "DocDate";
                    type = "F";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                {
                    errMessage = "담당자는 필수입니다.";
                    ClickCode = "CntcCode";
                    type = "F";
                    throw new Exception();
                }
                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    ClickCode = string.Empty;
                    type = string.Empty;
                    throw new Exception();
                }
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                    ClickCode = string.Empty;
                    type = string.Empty;
                    throw new Exception();
                }
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작업지시문서는 필수입니다.";
                        ClickCode = "PP030No";
                        type = "M";
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고코드는 필수입니다.";
                        ClickCode = "WhsCode";
                        type = "M";
                        throw new Exception();
                    }
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목코드는 필수입니다.";
                        ClickCode = "ItemCode";
                        type = "M";
                        throw new Exception();
                    }
                    if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "생산수량은 필수입니다.";
                        ClickCode = "PQty";
                        type = "M";
                        throw new Exception();
                    }

                    PP030HNo = oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value;
                    ORDRNo = oMat01.Columns.Item("ORDRNo").Cells.Item(i).Specific.Value;
                    RDR1No = oMat01.Columns.Item("RDR1No").Cells.Item(i).Specific.Value;
                    Query01 = "EXEC PS_PP081_02 '" + BPLID + "','" + OrdGbn + "','" + PP030HNo + "','" + ORDRNo + "','" + RDR1No + "'";
                    oRecordSet01.DoQuery(Query01);

                    if (Convert.ToDouble(oRecordSet01.Fields.Item("Qty").Value) < Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "수주잔량 수량보다 초과 생산할 수 없습니다.";
                        ClickCode = "PQty";
                        type = "M";
                        throw new Exception();
                    }

                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104")
                    {
                        if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value) == "Y")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.SetStatusBarMessage("부품,멀티작업은 배치번호가 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                functionReturnValue = false;
                                return functionReturnValue;
                            }
                        }
                    }
                    if (PS_PP081_Validate("검사01") == false)
                    {
                        functionReturnValue = false;
                        return functionReturnValue; //체크필요.
                    }
                }
                oDS_PS_PP081L.RemoveRecord(oDS_PS_PP081L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP081_FormClear();
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
                    else if(type == "M")
                    {
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// PS_PP081_LoadData
        /// </summary>
        private void PS_PP081_LoadData()
        {
            short i;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {                
                sQry = "EXEC [PS_PP081_01] '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "','" + oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() + "'";
                oRecordSet01.DoQuery(sQry);

                oMat02.Clear();
                oDS_PS_PP0811L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oForm.Freeze(true);

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP0811L.Size)
                    {
                        oDS_PS_PP0811L.InsertRecord(i);
                    }

                    oMat02.AddRow();
                    oDS_PS_PP0811L.Offset = i;
                    oDS_PS_PP0811L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("PP030No").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("Lotno").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("DocDate").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("DocDueDate").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("OrdNum").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("jqty").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("oqty").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("DocNum").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("PP030HNo").Value.ToString().Trim());
                    oDS_PS_PP0811L.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("PP030MNo").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
        /// 네비게이션 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_RECORD_MOVE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string Query01;
            string docEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                docEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim(); //현재문서번호

                if (pVal.MenuUID == "1288") //다음
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                        return;
                    }
                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                            return;
                        }
                    }
                    else
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("DocEntry").Enabled = true;
                        Query01 = "  SELECT ISNULL( MIN(DocEntry), (SELECT MIN(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('111','601')))";
                        Query01 += "   FROM [@PS_PP080H]";
                        Query01 += " WHERE U_OrdGbn IN ('102','602')";
                        Query01 += " AND DocEntry > " + docEntry;

                        oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(Query01, 0, 1);
                        oForm.Items.Item("1").Enabled = true;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                    }
                }
                else if (pVal.MenuUID == "1289") //이전
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                        return;
                    }
                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            return;
                        }
                    }
                    else
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        oForm.Items.Item("DocEntry").Enabled = true;
                        Query01 = "  SELECT ISNULL(MAX(DocEntry),(SELECT MAX(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('111','601')))";
                        Query01 += "   FROM [@PS_PP080H]";
                        Query01 += " WHERE U_OrdGbn IN ('102','602')";
                        Query01 += " AND DocEntry < " + docEntry;

                        oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(Query01, 0, 1);
                        oForm.Items.Item("1").Enabled = true;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                    }
                }
                else if (pVal.MenuUID == "1290") //최초
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    Query01 = " SELECT MIN(DocEntry)";
                    Query01 += "  FROM [@PS_PP080H]";
                    Query01 += " WHERE U_OrdGbn IN ('102','602')";

                    oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(Query01, 0, 1);
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                }
                else if (pVal.MenuUID == "1291") //최종
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    Query01 = " SELECT MAX(DocEntry)";
                    Query01 += "  FROM [@PS_PP080H]";
                    Query01 += " WHERE U_OrdGbn IN ('102','602')";

                    oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(Query01, 0, 1);
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                BubbleEvent = false;
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
                        PS_PP040 PS_PP040 = new PS_PP040();
                        PS_PP040.LoadForm(oForm.Items.Item("PP040No").Specific.Value);
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP081_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //Addon만등록 시 주석_S
                            if (PS_PP081_Add_InventoryGenEntry() == false)
                            {
                                PS_PP081_AddMatrixRow(oMat01.VisualRowCount, false);
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            //Addon만등록 시 주석_E
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP081_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
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
                                dataHelpClass.DoQuery("EXEC PS_PP080_03 '" + oDocEntry + "'");
                                PS_PP081_FormItemEnabled();
                                PS_PP081_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP081_FormItemEnabled();
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
                    if (pVal.ColUID == "PP030No")
                    {
                        if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
                        {
                            dataHelpClass.MDC_GF_Message("사업장은 필수입니다.", "W");
                            BubbleEvent = false;
                            return;
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                        {
                            dataHelpClass.MDC_GF_Message("작업구분은 필수입니다.", "W");
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP030No"); //사용자값활성
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP081_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else
                            {
                                oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP081H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP081_AddMatrixRow(0, true);
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP081_AddMatrixRow(0, true);
                                if (oForm.Items.Item("OrdGbn").Specific.Value == "102" || oForm.Items.Item("OrdGbn").Specific.Value == "602")
                                {
                                    PS_PP081_LoadData(); //부품포장대기 자료 SELECT
                                }
                                else
                                {
                                    oMat02.Clear();
                                    oMat02.FlushToDataSource();
                                    oMat02.LoadFromDataSource();
                                }
                            }
                            else
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
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
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                        }
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
        /// DOUBLE CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            int j;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat02" && pVal.Row != 0 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && !string.IsNullOrEmpty(oDS_PS_PP0811L.GetValue("U_ColReg05", pVal.Row - 1).ToString().Trim()))
                    {
                        if (oMat01.VisualRowCount == 0)
                        {
                            oDS_PS_PP081L.Clear();
                        }
                        j = 0;
                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (oDS_PS_PP081L.GetValue("U_PP030No", i).ToString().Trim() == oDS_PS_PP0811L.GetValue("U_ColReg01", pVal.Row - 1).ToString().Trim() && oDS_PS_PP081L.GetValue("U_BatchNum", i).ToString().Trim() == oDS_PS_PP0811L.GetValue("U_ColReg03", pVal.Row - 1).ToString().Trim())
                            {
                                dataHelpClass.MDC_GF_Message( "같은 행을 두번 선택할 수 없습니다. 확인하세요.", "W");
                                j = 1;
                            }
                        }
                        if (j == 0)
                        {
                            oDS_PS_PP081L.SetValue("U_PP030No", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg01", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_OrdGbn", oMat01.VisualRowCount - 1, oForm.Items.Item("OrdGbn").Specific.Value);
                            oDS_PS_PP081L.SetValue("U_OrdNum", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg05", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_OrdSub1", oMat01.VisualRowCount - 1, "00");
                            oDS_PS_PP081L.SetValue("U_OrdSub2", oMat01.VisualRowCount - 1, "000");
                            oDS_PS_PP081L.SetValue("U_ORDRNo", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg06", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_RDR1No", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg07", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_BPLId", oMat01.VisualRowCount - 1, oForm.Items.Item("BPLId").Specific.Value);
                            oDS_PS_PP081L.SetValue("U_PP030HNo", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg01", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_PP030MNo", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg09", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_ItemCode", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg04", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_ItemName", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg02", pVal.Row - 1).ToString().Trim());
                            oDS_PS_PP081L.SetValue("U_WhsCode", oMat01.VisualRowCount - 1, "101");
                            oDS_PS_PP081L.SetValue("U_WhsName", oMat01.VisualRowCount - 1, "창원");
                            oDS_PS_PP081L.SetValue("U_BatchNum", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColReg03", pVal.Row - 1).ToString().Trim());

                            if (Convert.ToDouble(oDS_PS_PP0811L.GetValue("U_ColQty01", pVal.Row - 1).ToString().Trim()) > Convert.ToDouble(oDS_PS_PP0811L.GetValue("U_ColQty02", pVal.Row - 1)))
                            {
                                oDS_PS_PP081L.SetValue("U_BQty", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColQty02", pVal.Row - 1).ToString().Trim());
                                oDS_PS_PP081L.SetValue("U_BWeight", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColQty02", pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                oDS_PS_PP081L.SetValue("U_BQty", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColQty01", pVal.Row - 1).ToString().Trim());
                                oDS_PS_PP081L.SetValue("U_BWeight", oMat01.VisualRowCount - 1, oDS_PS_PP0811L.GetValue("U_ColQty01", pVal.Row - 1).ToString().Trim());
                            }
                            oDS_PS_PP081L.SetValue("U_LineId", oMat01.VisualRowCount - 1, Convert.ToString(oMat01.VisualRowCount));

                            PS_PP081_AddMatrixRow(oMat01.VisualRowCount, false);

                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            j = 0;
                        }
                        BubbleEvent = false;
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
            double SumQty = 0;
            double Weight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PP030No")
                            {
                                for (i = 1; i <= oMat01.RowCount; i++)
                                {
                                    if (oMat01.Columns.Item("PP030No").Cells.Item(i).Specific.Value == oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                    {
                                        dataHelpClass.MDC_GF_Message("이미 입력한 작업지시문서입니다.", "W");
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP081_AddMatrixRow(pVal.Row, false);
                                            return;   
                                        }
                                    }
                                }
                                Query01 = "EXEC PS_PP081_02 '" + oMat01.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value + "','" + oForm.Items.Item("OrdGbn").Specific.Selected.Value + "'";
                                oRecordSet01.DoQuery(Query01);
                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                }
                                else
                                {
                                    oDS_PS_PP081L.SetValue("U_PP030No", pVal.Row - 1, oRecordSet01.Fields.Item("PP030No").Value);
                                    oDS_PS_PP081L.SetValue("U_OrdGbn", pVal.Row - 1, oRecordSet01.Fields.Item("OrdGbn").Value);
                                    oDS_PS_PP081L.SetValue("U_OrdNum", pVal.Row - 1, oRecordSet01.Fields.Item("OrdNum").Value);
                                    oDS_PS_PP081L.SetValue("U_OrdSub1", pVal.Row - 1, oRecordSet01.Fields.Item("OrdSub1").Value);
                                    oDS_PS_PP081L.SetValue("U_OrdSub2", pVal.Row - 1, oRecordSet01.Fields.Item("OrdSub2").Value);
                                    oDS_PS_PP081L.SetValue("U_PP030HNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030HNo").Value);
                                    oDS_PS_PP081L.SetValue("U_PP030MNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030MNo").Value);
                                    oDS_PS_PP081L.SetValue("U_ORDRNo", pVal.Row - 1, oRecordSet01.Fields.Item("ORDRNo").Value);
                                    oDS_PS_PP081L.SetValue("U_RDR1No", pVal.Row - 1, oRecordSet01.Fields.Item("RDR1No").Value);
                                    oDS_PS_PP081L.SetValue("U_BPLId", pVal.Row - 1, oRecordSet01.Fields.Item("BPLId").Value);
                                    oDS_PS_PP081L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP081L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP081L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet01.Fields.Item("CpCode").Value);
                                    oDS_PS_PP081L.SetValue("U_CpName", pVal.Row - 1, oRecordSet01.Fields.Item("CpName").Value);
                                    oDS_PS_PP081L.SetValue("U_BQty", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("BQty").Value));
                                    oDS_PS_PP081L.SetValue("U_BWeight", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("BWeight").Value));
                                    oDS_PS_PP081L.SetValue("U_PQty", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("PQty").Value));
                                    oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("PWeight").Value));
                                    oDS_PS_PP081L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("YQty").Value));
                                    oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("YWeight").Value));
                                    oDS_PS_PP081L.SetValue("U_NQty", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("NQty").Value));
                                    oDS_PS_PP081L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("NWeight").Value));
                                    oDS_PS_PP081L.SetValue("U_WhsCode", pVal.Row - 1, oRecordSet01.Fields.Item("WhsCode").Value);
                                    oDS_PS_PP081L.SetValue("U_WhsName", pVal.Row - 1, oRecordSet01.Fields.Item("WhsName").Value);
                                    oDS_PS_PP081L.SetValue("U_BatchNum", pVal.Row - 1, oRecordSet01.Fields.Item("BatchNum").Value);
                                    oDS_PS_PP081L.SetValue("U_LineId", pVal.Row - 1, oRecordSet01.Fields.Item("LineId").Value);

                                    oMat01.LoadFromDataSource();

                                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)//합격수량 sum
                                    {
                                        SumQty += oMat01.Columns.Item("YQty").Cells.Item(i + 1).Specific.Value;
                                    }
                                    oForm.Items.Item("SumQty").Specific.Value = SumQty;
                                }
                            }
                            else if (pVal.ColUID == "PQty")
                            {
                                if (oMat01.Columns.Item("OrdGbn").Cells.Item(pVal.Row).Specific.Value == "102")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value) < Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        //포장기준 수량보다 많이 입고잡을수 없다.
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item("BQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value));
                                        oDS_PS_PP081L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value));
                                        Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                        if (Weight == 0)
                                        {
                                            oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else
                                        {
                                            oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        oDS_PS_PP081L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP081L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else
                                    {
                                        oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        oDS_PS_PP081L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                        if (Weight == 0)
                                        {
                                            oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else
                                        {
                                            oDS_PS_PP081L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        oDS_PS_PP081L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP081L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }
                                oMat01.LoadFromDataSourceEx();

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)//합격수량 sum
                                {
                                    SumQty += Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i + 1).Specific.Value);
                                }
                                oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            }
                            else if (pVal.ColUID == "NQty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                }
                                else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP081L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                }
                                else
                                {
                                    oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP081L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                    if (Weight == 0)
                                    {
                                        oDS_PS_PP081L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else
                                    {
                                        oDS_PS_PP081L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        oDS_PS_PP081L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                    }
                                }
                                oMat01.LoadFromDataSourceEx();

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++) //합격수량 sum
                                {
                                    SumQty += Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i + 1).Specific.Value);
                                }
                                oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            }
                            else
                            {
                                oDS_PS_PP081L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP081H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CardCode")
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP081H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP081H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP081H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            double sumQty = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value))
                            {
                                sumQty += 0;
                            }
                            else
                            {
                                sumQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value);
                            }
                        }
                        oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(sumQty);

                        PS_PP081_AddMatrixRow(oMat01.VisualRowCount, false);
                    }
                    PS_PP081_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP081H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP081L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP0811L);
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
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PS_PP081_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST( 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
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
                                oDS_PS_PP081L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_PP081L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if(oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
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
            double sumQty = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_PP081_Validate("행삭제") == false)
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
                        oDS_PS_PP081L.RemoveRecord(oDS_PS_PP081L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP081_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP081L.GetValue("U_PP030No", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP081_AddMatrixRow(oMat01.RowCount, false);
                            }
                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++) //합격수량 sum
                            {
                                sumQty += Convert.ToDouble(oMat01.Columns.Item("YQty").Cells.Item(i + 1).Specific.Value);
                            }
                            oForm.Items.Item("SumQty").Specific.Value = sumQty;
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
            string errMessage = string.Empty;
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
                            errMessage = "현재화면에서 취소할수 없습니다. 생산완료등록(PP080) 화면에서 취소하시기 바랍니다.";
                            BubbleEvent = false;
                            throw new Exception();
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
                            Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":  //취소
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            PS_PP081_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Specific.Value = oDocEntry;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP081_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP081_FormItemEnabled();
                            PS_PP081_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
                            PS_PP081_FormItemEnabled();
                            PS_PP081_AddMatrixRow(0, true);
                            break;
                    }
                }
                return;
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
