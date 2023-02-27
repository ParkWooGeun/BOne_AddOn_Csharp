using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작업일보등록(Powder)
    /// </summary>
    internal class PS_PP052 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.Matrix oMat04;
        private SAPbouiCOM.DBDataSource oDS_PS_PP052H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP052L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP052M;
        private SAPbouiCOM.DBDataSource oDS_PS_PP052N;
        private SAPbouiCOM.DBDataSource oDS_PS_PP052W;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private int oMat01Row01;
        private int oMat02Row02;
        private int oMat03Row03;

        private string oDocEntry01;
        private string oOrdGbn;
        private string oSequence;
        private string oDocdate;
        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP052.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP052_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP052");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP052_CreateItems();
                PS_PP052_ComboBox_Setting();
                PS_PP052_CF_ChooseFromList();
                PS_PP052_EnableMenus();
                PS_PP052_SetDocument(oFormDocEntry);
                PS_PP052_FormResize();
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
        private void PS_PP052_CreateItems()
        {
            try
            {
                oDS_PS_PP052H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
                oDS_PS_PP052L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
                oDS_PS_PP052M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
                oDS_PS_PP052N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");
                oDS_PS_PP052W = oForm.DataSources.DBDataSources.Item("@PS_PP040W");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oMat04 = oForm.Items.Item("Mat04").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);

                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");
                oForm.Items.Item("Opt04").Specific.DataBind.SetBound(true, "", "Opt04");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt04");

                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

                oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");

                oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP052_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "10", "일반");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP040", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP040", "DocType", false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE IN('601','111') order by Code", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");

                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_PP052_CF_ChooseFromList()
        {
            ChooseFromListCollection oCFLs = null;
            Conditions oCons = null;
            Condition oCon = null;
            ChooseFromList oCFL = null;
            ChooseFromListCreationParams oCFLCreationParams = null;
            EditText oEdit = null;

            try
            {
                oEdit = oForm.Items.Item("ItemCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLITEMCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "ItmsGrpCod";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "102";
                oCFL.SetConditions(oCons);

                oEdit.ChooseFromListUID = "CFLITEMCODE";
                oEdit.ChooseFromListAlias = "ItemCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }
                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }
                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }
            }
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP052_Validate(string ValidateType)
        {
            bool returnValue = false;
            int i = 0;
            int j;
            string Query01;
            string LineNum;
            string DocEntry;
            string type = string.Empty;
            string ClickCode = string.Empty;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.";
                        throw new Exception();
                    }
                }
                if (ValidateType == "검사01") //작업타입이 일반인경우
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10")
                    {
                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //입력된 행에 대해
                        {
                            if (Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "작업지시문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }

                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)//삭제된 행에 대한처리
                        {
                            Query01 =  "SELECT ";
                            Query01 += " PS_PP040H.DocEntry,";
                            Query01 += " PS_PP040L.LineId,";
                            Query01 += " CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                            Query01 += " PS_PP040L.U_OrdGbn AS OrdGbn,";
                            Query01 += " PS_PP040L.U_PP030HNo AS PP030HNo,";
                            Query01 += " PS_PP040L.U_PP030MNo AS PP030MNo,";
                            Query01 += " PS_PP040L.U_OrdMgNum AS OrdMgNum ";
                            Query01 += " FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                            Query01 += " WHERE PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";

                            oRecordSet01.DoQuery(Query01);
                            for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                            {
                                bool Exist = false;
                                //기존에 있는 행에대한처리
                                for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                {
                                    //새로추가된 행인경우, 검사할필요없다
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                    {
                                    }
                                    else
                                    {
                                        //라인번호가 같고, 문서번호가 같으면 존재하는행
                                        if (Convert.ToInt32(oRecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value) && Convert.ToInt32(oRecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                        {
                                            Exist = true;
                                        }
                                    }
                                }
                                oRecordSet01.MoveNext();
                            }
                        }
                    }
                }
                else if (ValidateType == "수정01")
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))
                    {
                    }
                    else
                    {
                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "111" || oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "601")
                        {
                            if (oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.Value == "CP80111" || oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.Value == "CP80101")
                            {
                                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                                LineNum = oMat01.Columns.Item("LineNum").Cells.Item(oMat01Row01).Specific.Value;

                                if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(oMat01Row01).Specific.Value) != Convert.ToDouble(dataHelpClass.GetValue("select U_pqty from [@PS_PP040L] where DocEntry ='" + DocEntry + "' and u_linenum ='" + LineNum + "'", 0, 1)))
                                {
                                    errMessage = "원자재 불출이 진행된 행은 생산수량을 수정할 수 없습니다.";
                                    oMat01.Columns.Item("PQty").Cells.Item(oMat01Row01).Specific.Value = Convert.ToDouble(dataHelpClass.GetValue("select U_pqty from [@PS_PP040L] where DocEntry ='" + DocEntry + "' and u_linenum ='" + LineNum + "'", 0, 1));
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errMessage = "이미취소된 문서 입니다. 취소할수 없습니다.";
                        throw new Exception();
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP052_EnableMenus()
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
        /// PS_PP052_SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP052_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP052_FormItemEnabled();
                    PS_PP052_AddMatrixRow01(0, true);
                    PS_PP052_AddMatrixRow02(0, true);
                    PS_PP052_AddMatrixRow04(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP052_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
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
        private void PS_PP052_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat02").Top = 189;
                oForm.Items.Item("Mat02").Left = 7;
                oForm.Items.Item("Mat02").Height = ((oForm.Height - 170) / 4 * 1) - 20;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 20;
                oMat02.AutoResizeColumns();

                oForm.Items.Item("Mat03").Top = 369;
                oForm.Items.Item("Mat03").Left = 7;
                oForm.Items.Item("Mat03").Height = ((oForm.Height - 170) / 4 * 1) - 20;
                oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 20;
                oMat03.AutoResizeColumns();

                oForm.Items.Item("Mat04").Top = 189;
                oForm.Items.Item("Mat04").Left = oForm.Width / 2;
                oForm.Items.Item("Mat04").Height = ((oForm.Height - 170) / 2 * 1);
                oForm.Items.Item("Mat04").Width = oForm.Width / 2 - 50;
                oMat04.AutoResizeColumns();

                oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 40;
                oForm.Items.Item("Mat01").Left = 7;
                oForm.Items.Item("Mat01").Height = ((oForm.Height - 170) / 4 * 2) - 160;
                oForm.Items.Item("Mat01").Width = oForm.Width - 50;
                oMat01.AutoResizeColumns();

                oForm.Items.Item("Opt01").Left = 10;
                oForm.Items.Item("Opt02").Left = 10;
                oForm.Items.Item("Opt04").Left = oForm.Width / 2;
                oForm.Items.Item("Opt03").Left = 10;
                oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;

                oForm.Items.Item("Button02").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP052_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oMat01.Columns.Item("Sequence").Visible = false;
                oMat01.Columns.Item("OrdGbn").Visible = false;
                oMat01.Columns.Item("BPLId").Visible = false;
                oMat01.Columns.Item("OrdNum").Visible = false;
                oMat01.Columns.Item("OrdSub1").Visible = false;
                oMat01.Columns.Item("OrdSub2").Visible = false;
                oMat01.Columns.Item("PP030HNo").Visible = false;
                oMat01.Columns.Item("PP030MNo").Visible = false;
                oMat01.Columns.Item("SelWt").Visible = false;
                oMat01.Columns.Item("PSum").Visible = false;
                oMat01.Columns.Item("BQty").Visible = false;
                oMat01.Columns.Item("SubLot").Visible = false;
                oMat01.Columns.Item("LineId").Visible = false;
                oMat01.Columns.Item("BdwQty").Visible = false;
                oMat01.Columns.Item("DwRate").Visible = false;
                oMat01.Columns.Item("AdwQty").Visible = false;
                oMat01.Columns.Item("NdwQTy").Visible = false;
                oMat01.Columns.Item("MachGrCd").Visible = false;
                oMat01.Columns.Item("CompltYN").Visible = false;
                oMat01.Columns.Item("WorkCls").Visible = false;
                oMat01.Columns.Item("SCpCode").Visible = false;
                oMat01.Columns.Item("SCpName").Visible = false;
                oMat01.Columns.Item("MachCode").Visible = false;
                oMat01.Columns.Item("MachName").Visible = false;
                oMat01.Columns.Item("YQty").Visible = false;
                oMat01.Columns.Item("YWeight").Visible = false;
                oMat01.Columns.Item("PWeight").Visible = false;
                oMat01.Columns.Item("NWeight").Visible = false;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
                oMat04.AutoResizeColumns();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oForm.Items.Item("Mat03").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat02.Columns.Item("NTime").Editable = true; //비가동시간만 사용
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oMat01.Columns.Item("OrdMgNum").Editable = true;
                    oMat01.Columns.Item("PQty").Editable = true;
                    oMat01.Columns.Item("NQty").Editable = true;
                    oMat01.Columns.Item("ScrapWt").Editable = true;
                    oMat01.Columns.Item("WorkTime").Editable = true;
                    oMat01.Columns.Item("CItemCod").Editable = true;
                    oForm.Items.Item("CpCode").Enabled = true;
                    oForm.Items.Item("BaseTime").Enabled = true;
                    oForm.Items.Item("WorkTime").Enabled = true;
                    oForm.Items.Item("Shift").Enabled = true;

                    if (string.IsNullOrEmpty(oOrdGbn))
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    PS_PP052_FormClear();

                    oDS_PS_PP052H.SetValue("U_DocType", 0, "10");

                    if (string.IsNullOrEmpty(oDocdate))
                    {
                        oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
                    }
                    else
                    {
                        oForm.Items.Item("DocDate").Specific.Value = oDocdate;
                    }
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Mat03").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdType").Enabled = false;
                    oForm.Items.Item("OrdMgNum").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("Button01").Enabled = false;
                    oForm.Items.Item("1").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Mat03").Enabled = false;
                    oForm.Items.Item("Mat04").Enabled = false;
                    oMat01.Columns.Item("OrdMgNum").Editable = false;
                    oMat01.Columns.Item("PQty").Editable = false;
                    oMat01.Columns.Item("NQty").Editable = false;
                    oMat01.Columns.Item("ScrapWt").Editable = false;
                    oMat01.Columns.Item("WorkTime").Editable = false;
                    oMat01.Columns.Item("CItemCod").Editable = false;
                    oForm.Items.Item("CpCode").Enabled = false;
                    oForm.Items.Item("BaseTime").Enabled = false;
                    oForm.Items.Item("WorkTime").Enabled = false;
                    oForm.Items.Item("OrdGbn").Enabled = false;
                    oForm.Items.Item("Shift").Enabled = false;
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
        /// PS_PP052_AddMatrixRow01
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP052_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP052L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP052L.Offset = oRow;
                oDS_PS_PP052L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP052L.SetValue("U_WorkCls", oRow, "A");
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
        /// PS_PP052_AddMatrixRow02
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP052_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP052M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP052M.Offset = oRow;
                oDS_PS_PP052M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
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
        /// PS_PP052_AddMatrixRow03
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP052_AddMatrixRow03(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP052N.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP052N.Offset = oRow;
                oDS_PS_PP052N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat03.LoadFromDataSource();
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
        /// PS_PP052_AddMatrixRow04
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP052_AddMatrixRow04(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP052W.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP052W.Offset = oRow;
                oDS_PS_PP052W.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat03.LoadFromDataSource();
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
        private void PS_PP052_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "");
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
        private bool PS_PP052_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            int j;
            double FailQty;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP052_FormClear();
                }
                if (Convert.ToDouble(dataHelpClass.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.Value + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'", 0, 1)) > 0)
                {
                    errMessage = "해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.";
                    type = "X";
                    throw new Exception();
                }
                else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                {
                    type = "X";
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                {
                    errMessage = "작지번호는 필수입니다.";
                    ClickCode = "OrdNum";
                    type = "F";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount == 1)
                {
                    type = "X";
                    errMessage = "공정정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                else if (oMat02.VisualRowCount == 1)
                {
                    type = "X";
                    errMessage = "작업자정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                else if (oMat03.VisualRowCount == 0)
                {
                    type = "X";
                    errMessage = "불량정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작지번호는 필수입니다.";
                        ClickCode = "OrdMgNum";
                        type = "M";
                        throw new Exception();
                    }
                    if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "생산수량은 필수입니다..";
                        ClickCode = "PQty";
                        type = "M";
                        throw new Exception();
                    }
                    if (Convert.ToDouble(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "실동시간은 필수입니다.";
                        ClickCode = "WorkTime";
                        type = "M";
                        throw new Exception();
                    }
                    if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80101" || oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80111")
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("Charge").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "용해 Charge 입력은 필수입니다.";
                            ClickCode = "Charge";
                            type = "M";
                            throw new Exception();
                        }
                    }
                    FailQty = 0;
                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 && string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
                        {
                            errMessage = "불량수량이 입력되었을 때는 불량코드는 필수입니다.";
                            ClickCode = "FailCode";
                            type = "M";
                            throw new Exception();
                        }
                        else if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) == 0 && !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value))
                        {
                            errMessage = "불량코드를 확인하세요.";
                            ClickCode = "FailCode";
                            type = "M";
                            throw new Exception();
                        }
                        if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value) && (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value))
                        {
                            FailQty += Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value);
                        }
                    }
                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value) == 1 && string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요.";
                            type = "X";
                            throw new Exception();
                        }
                    }
                }

                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동코드가 입력되었을 때는 비가동시간은 필수입니다.";
                            ClickCode = "NTime";
                            type = "M";
                            throw new Exception();
                        }
                    }
                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동시간이 입력되었을 때는 비가동코드는 필수입니다.";
                            ClickCode = "NCode";
                            type = "M";
                            throw new Exception();
                        }
                    }
                }

                if (PS_PP052_Validate("검사01") == false)
                {
                    returnValue = false;
                    return returnValue;
                }
                oDS_PS_PP052L.RemoveRecord(oDS_PS_PP052L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP052M.RemoveRecord(oDS_PS_PP052M.Size - 1);
                oMat02.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP052_FormClear();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_PP052_SumWorkTime
        /// </summary>
        private void PS_PP052_SumWorkTime()
        {
            int loopCount;
            double Total = 0;

            try
            {
                for (loopCount = 0; loopCount <= oMat01.RowCount - 2; loopCount++)
                {
                    Total += Convert.ToDouble((string.IsNullOrEmpty(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value.ToString().Trim()) ? 0 : oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value.ToString().Trim()));
                }
                oForm.Items.Item("Total").Specific.Value = Total.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PP052_FindValidateDocument
        /// </summary>
        private bool PS_PP052_FindValidateDocument(string ObjectType)
        {
            bool returnValue = false;
            string Query01;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                Query01 = " SELECT DocEntry";
                Query01 += " FROM [" + ObjectType + "] Where DocEntry = ";
                Query01 += DocEntry;
                Query01 += " AND U_DocType = '10'";

                oRecordSet01.DoQuery(Query01);
                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.";
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_PP052_OrderInfoLoad
        /// </summary>
        private void PS_PP052_OrderInfoLoad()
        {
            string Query01;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                {
                    errMessage = "작업지시 관리번호를 입력하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    Query01 = "SELECT ";
                    Query01 += "U_OrdGbn,";
                    Query01 += "U_BPLId,";
                    Query01 += "U_ItemCode,";
                    Query01 += "U_ItemName,";
                    Query01 += "U_OrdNum,";
                    Query01 += "U_OrdSub1,";
                    Query01 += "U_OrdSub2,";
                    Query01 += "DocEntry";
                    Query01 += " FROM [@PS_PP030H]";
                    Query01 += " WHERE ";
                    Query01 += " U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'";
                    Query01 += " AND U_OrdGbn NOT IN('104','107') ";
                    Query01 += " AND Canceled = 'N'";
                    oRecordSet01.DoQuery(Query01);

                    if (oRecordSet01.RecordCount == 0)
                    {
                        errMessage = "작업지시 정보가 존재하지 않습니다.";
                        throw new Exception();
                    }
                    else
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(oRecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("BPLId").Specific.Select(oRecordSet01.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("ItemCode").Specific.Value = oRecordSet01.Fields.Item(2).Value;
                        oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(3).Value;
                        oForm.Items.Item("OrdNum").Specific.Value = oRecordSet01.Fields.Item(4).Value;
                        oForm.Items.Item("OrdSub1").Specific.Value = oRecordSet01.Fields.Item(5).Value;
                        oForm.Items.Item("OrdSub2").Specific.Value = oRecordSet01.Fields.Item(6).Value;
                        oForm.Items.Item("PP030HNo").Specific.Value = oRecordSet01.Fields.Item(7).Value;
                        oForm.Update();
                    }
                    dataHelpClass.Set_ComboList(oForm.Items.Item("CpCode").Specific, "select U_CpCode as Code ,U_CpName as Name  from [@PS_PP004H] where U_ItemCode ='" + oRecordSet01.Fields.Item(2).Value + "' and U_CpCode not in ('CP80198','CP80199')", "", true, false);
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_PP052_Add_InventoryGenExit
        /// </summary>
        /// <returns></returns>
        private bool PS_PP052_Add_InventoryGenExit()
        {
            bool returnValue = false;
            int i;
            int j = 0;
            int RetVal;
            int Cnt;
            int errDiCode = 0;
            string sQry;
            string sQry02;
            string SDocEntry;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            double Price;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClss = new PSH_CodeHelpClass();
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.TaxDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.Comments = "원재료 불출 등록(" + oDS_PS_PP052H.GetValue("DocEntry", 0).ToString().Trim() + ") 출고 - PS_PP052 ";

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    if ((oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP80101" || oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP80111") && !string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value) && Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) >= 0 && Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value) != 0)
                    {
                        sQry = "  SELECT PRICE";
                        sQry += "   FROM OIVL a INNER JOIN OIGN b ON a.BASE_REF = b.DocEntry AND b.U_Comments = 'Convert Meterial'";
                        sQry += "  WHERE a.ITEMCODE = '" + oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value + "'"; //CItemCod
                        sQry += "   AND CONVERT(CHAR(6), a.DocDate,112) = '" + codeHelpClss.Left(oForm.Items.Item("DocDate").Specific.Value, 6) + "'";

                        oRecordSet01.DoQuery(sQry);

                        Price = Convert.ToString(oRecordSet01.Fields.Item(0).Value) == "" ? 0 : Convert.ToDouble(oRecordSet01.Fields.Item(0).Value);

                        sQry02 = "select U_ItmBsort  from OITM where ItemCode ='" + oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                        oRecordSet02.DoQuery(sQry02);

                        oDIObject.Lines.Add();
                        oDIObject.Lines.SetCurrentLine(j);
                        oDIObject.Lines.ItemCode = oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value;
                        oDIObject.Lines.WarehouseCode = "101";
                        oDIObject.Lines.Quantity = float.Parse(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value);
                        oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value;

                        if (oRecordSet02.Fields.Item(0).Value == "501") //스크랩 품목일 경우 단가,금액 0
                        {
                            oDIObject.Lines.Price = 0;
                            oDIObject.Lines.UnitPrice = 0;
                            oDIObject.Lines.LineTotal = 0;
                        }
                        else if (oRecordSet01.RecordCount > 0) //제품원재료 변환 품목은 단가를 계산 후 입력
                        {
                            oDIObject.Lines.Price = Price;
                            oDIObject.Lines.UnitPrice = Price;
                            oDIObject.Lines.LineTotal = Price * Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value);
                        }

                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value == "111")
                        {
                            if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value) == "Y")
                            {
                                oDIObject.Lines.BatchNumbers.BatchNumber = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                                oDIObject.Lines.BatchNumbers.Quantity = float.Parse(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value);
                                oDIObject.Lines.BatchNumbers.Add();
                            }
                        }
                        j += 1;
                    }
                }
                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out SDocEntry);

                    Cnt = 1;
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value == "CP80101" || oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value == "CP80111")
                        {
                            oDS_PS_PP052L.SetValue("U_OutDoc", i, SDocEntry);
                            oDS_PS_PP052L.SetValue("U_OutLin", i, Convert.ToString(Cnt));
                            Cnt += 1;
                        }
                    }
                }

                oMat01.LoadFromDataSource();
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                oMat01.AutoResizeColumns();
                returnValue = true;
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
        /// PS_PP052_Add_InventoryGenEntry
        /// </summary>
        /// <returns></returns>
        private bool PS_PP052_Add_InventoryGenEntry()
        {
            bool returnValue = true;
            int i;
            int j = 0;
            int RetVal;
            int errDiCode = 0;
            string sQry;
            string sQry02;
            string SDocEntry;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            double Price;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClss = new PSH_CodeHelpClass();
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();
                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.TaxDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.Value.ToString().Trim();
                oDIObject.Comments = "원재료 불출 등록 취소(" + oDS_PS_PP052H.GetValue("DocEntry", 0).ToString().Trim() + ") 입고 - PS_PP052 ";

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    if ((oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP80101" || oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP80111") && !string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value) && Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) >= 0 && Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value) != 0)
                    {
                        sQry = "  SELECT PRICE";
                        sQry += "   FROM OIVL a INNER JOIN OIGN b ON a.BASE_REF = b.DocEntry AND b.U_Comments = 'Convert Meterial'";
                        sQry += "  WHERE a.ITEMCODE = '" + oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value + "'";
                        sQry += "   AND CONVERT(CHAR(6), a.DocDate,112) = '" + codeHelpClss.Left(oForm.Items.Item("DocDate").Specific.Value, 6) + "'";

                        oRecordSet01.DoQuery(sQry);

                        Price = Convert.ToString(oRecordSet01.Fields.Item(0).Value) == "" ? 0 : Convert.ToDouble(oRecordSet01.Fields.Item(0).Value);

                        sQry02 = "select U_ItmBsort  from OITM where ItemCode ='" + oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                        oRecordSet02.DoQuery(sQry02);

                        oDIObject.Lines.Add();
                        oDIObject.Lines.SetCurrentLine(j);
                        oDIObject.Lines.ItemCode = oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value;
                        oDIObject.Lines.WarehouseCode = "101";
                        oDIObject.Lines.Quantity = float.Parse(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value);
                        oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value;

                        if (oRecordSet02.Fields.Item(0).Value == "501") //스크랩 품목일 경우 단가,금액 0
                        {
                            oDIObject.Lines.Price = 0;
                            oDIObject.Lines.UnitPrice = 0;
                            oDIObject.Lines.LineTotal = 0;
                        }
                        else if (oRecordSet01.RecordCount > 0) //제품원재료 변환 품목은 단가를 계산 후 입력
                        {
                            oDIObject.Lines.Price = Price;
                            oDIObject.Lines.UnitPrice = Price;
                            oDIObject.Lines.LineTotal = Price * Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i).Specific.Value);
                        }

                        if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Selected.Value == "111")
                        {
                            if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value) == "Y")
                            {
                                oDIObject.Lines.BatchNumbers.BatchNumber = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
                                oDIObject.Lines.BatchNumbers.Quantity = float.Parse(oMat01.Columns.Item("YQty").Cells.Item(i).Specific.Value);
                                oDIObject.Lines.BatchNumbers.Add();
                            }
                        }
                        j += 1;
                    }
                }
                RetVal = oDIObject.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out SDocEntry);

                    sQry = "Update [@PS_PP040L] set U_OutDocC = '" + SDocEntry + "', U_OutLinC = U_OutLin";
                    sQry = sQry + " From [@PS_PP040L] where 1=1 and u_cpcode in ('CP80101','CP80111') and docentry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' ";
                    oRecordSet01.DoQuery(sQry);
                }

                oMat01.LoadFromDataSource();
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                oMat01.AutoResizeColumns();
                returnValue = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            double tot_time;
            double UnitTime;
            double UnitRemainTime;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP052_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80101" || oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80111")
                            {
                                if (PS_PP052_Add_InventoryGenExit() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oOrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                            oSequence = oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.Value;
                            oDocdate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP052_DataValidCheck() == false)
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

                    if (pVal.ItemUID == "2")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oMat01.VisualRowCount > 1)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소") == 2)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP052_OrderInfoLoad();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            PS_PP052_OrderInfoLoad();
                        }
                    }
                    if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("WorkTime").Specific.Value))
                            {
                                errMessage = "오류 : 작업시간을 입력하세요.";
                                throw new Exception();
                            }
                            tot_time = Convert.ToDouble(oForm.Items.Item("WorkTime").Specific.Value);

                            if (tot_time > 0)
                            {
                                UnitTime = Convert.ToDouble(String.Format("{0:0.##}", (tot_time / (oMat01.VisualRowCount - 1))));
                                UnitRemainTime = Convert.ToDouble(String.Format("{0:0.##}", tot_time - UnitTime * (oMat01.VisualRowCount - 1)));

                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    if (i != oMat01.VisualRowCount - 2)
                                    {
                                        oDS_PS_PP052L.SetValue("U_WorkTime", i, Convert.ToString(UnitTime));
                                    }
                                    else
                                    {
                                        oDS_PS_PP052L.SetValue("U_WorkTime", i, Convert.ToString(UnitTime + UnitRemainTime));
                                    }
                                }
                            }
                            oMat01.LoadFromDataSource();
                            PS_PP052_SumWorkTime();

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
                                PS_PP052_FormItemEnabled();
                                PS_PP052_AddMatrixRow01(0, true);
                                PS_PP052_AddMatrixRow02(0, true);
                                PS_PP052_AddMatrixRow04(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PS_PP052_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP052_FormItemEnabled();
                            }
                        }
                    }
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
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "OrdMgNum")
                    {
                         dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "OrdMgNum", ""); //사용자값활성
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                            {
                                errMessage = "작업구분이 선택되지 않았습니다.";
                                throw new Exception();
                            }
                            else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
                            {
                                errMessage = "사업장이 선택되지 않았습니다.";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                            {
                                errMessage = "품목코드가 선택되지 않았습니다.";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                            {
                                errMessage = "작지번호가 선택되지 않았습니다.";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oForm.Items.Item("PP030HNo").Specific.Value))
                            {
                                errMessage = "작지문서번호가 선택되지 않았습니다.";
                                throw new Exception();
                            }
                            else
                            {
                                dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum"); //사용자값활성
                            }
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "WorkCode")
                        {
                            if (Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) == 0)
                            {
                                errMessage = "기준시간을 입력하지 않았습니다.";
                                type = "F";
                                ClickCode = "BaseTime";
                                throw new Exception();
                            }
                            if (string.IsNullOrEmpty(oForm.Items.Item("Shift").Specific.Value.ToString().Trim()))
                            {
                                errMessage = "주야 구분을 입력하지 않았습니다.";
                                type = "F";
                                ClickCode = "Shift";
                                throw new Exception();
                            }
                            if (string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value.ToString().Trim()))
                            {
                                errMessage = "공정코드를 입력하지 않았습니다.";
                                type = "F";
                                ClickCode = "CpCode";
                                throw new Exception();
                            }
                        }
                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "WorkCode"); //사용자값활성
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "NCode"); //사용자값활성
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "FailCode"); //사용자값활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode"); //설비코드 사용자값활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CItemCod"); //원재료코드 사용자값활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SCpCode"); //지원공정추가(2018.05.30 송명규)
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UseMCode", ""); //작업장비 사용자값활성
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //위쪽 화살표
                        {
                            if (pVal.Row > 1 && pVal.Row <= oMat01.VisualRowCount)
                            {
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        else if (pVal.CharPressed == 40) //아래 화살표
                        {
                            if (pVal.Row > 0 && pVal.Row < oMat01.VisualRowCount)
                            {
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        if (pVal.ColUID == "WorkTime" && pVal.Row != Convert.ToDouble("0")) //작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
                        {
                            PS_PP052_SumWorkTime();
                        }
                    }
                    else if (pVal.ItemUID == "WorkTime")
                    {
                        if (pVal.CharPressed == 9) //탭 키 Press
                        {
                            oMat02.Columns.Item("WorkCode").Cells.Item(1).Click();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
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
            finally
            {
                oForm.Freeze(false);
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
                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oMat01Row01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat02")
                {
                    if (pVal.Row > 0)
                    {
                        oMat02Row02 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat03")
                {
                    if (pVal.Row > 0)
                    {
                        oMat03Row03 = pVal.Row;
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                            else
                            {
                                oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                            else
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID != "특정컬럼")
                            {
                                oDS_PS_PP052N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat04")
                        {
                            if (pVal.ColUID != "특정컬럼")
                            {
                                oDS_PS_PP052N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "CpCode")
                            {
                                if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() != "CP80101" && oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() != "CP80111")
                                {
                                    oMat01.Columns.Item("Charge").Editable = false;
                                }
                                else
                                {
                                    oMat01.Columns.Item("Charge").Editable = true;
                                }
                                sQry = "SELECT U_Sequence FROM [@PS_PP030M] WHERE DocEntry = '" + oForm.Items.Item("PP030HNo").Specific.Value.ToString().Trim() + "' and U_CpCode = '" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'";

                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item(0).Value == 1)
                                {
                                    oMat01.Columns.Item("CItemCod").Editable = true;
                                }
                                else
                                {
                                    oMat01.Columns.Item("CItemCod").Editable = false;
                                }
                                if (Convert.ToInt32(DateTime.Now.ToString("HHmm")) > 830  && Convert.ToInt32(DateTime.Now.ToString("HHmm")) < 2030)
                                {
                                    oForm.Items.Item("Shift").Specific.Select("D", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    oForm.Items.Item("Shift").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                sQry = "select U_Minor, U_CdName from [@PS_SY001L] where code ='P216' and U_RelCd ='" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "' ";
                                oRecordSet01.DoQuery(sQry);

                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                                {
                                    if (i != 0)
                                    {
                                        oDS_PS_PP052W.InsertRecord(i);
                                    }
                                    oDS_PS_PP052W.Offset = i;
                                    oDS_PS_PP052W.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                    oDS_PS_PP052W.SetValue("U_WLCode", i, oRecordSet01.Fields.Item(0).Value);
                                    oDS_PS_PP052W.SetValue("U_WorkList", i, oRecordSet01.Fields.Item(1).Value);

                                    oRecordSet01.MoveNext();
                                }
                                oMat04.LoadFromDataSource();

                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP052_AddMatrixRow01(0, true);

                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP052_AddMatrixRow02(0, true);

                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }

                            if (pVal.ItemUID == "OrdType")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oForm.Items.Item("OrdGbn").Enabled = true;
                                oForm.Items.Item("BPLId").Enabled = false;
                                oForm.Items.Item("ItemCode").Enabled = false;
                                oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                oForm.Items.Item("ItemCode").Specific.Value = "";
                                oForm.Items.Item("ItemName").Specific.Value = "";
                                oForm.Items.Item("OrdNum").Specific.Value = "";
                                oForm.Items.Item("OrdSub1").Specific.Value = "";
                                oForm.Items.Item("OrdSub2").Specific.Value = "";
                                oForm.Items.Item("PP030HNo").Specific.Value = "";
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP052_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP052_AddMatrixRow02(0, true);
                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                oMat04.LoadFromDataSource();
                                PS_PP052_AddMatrixRow04(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                                dataHelpClass.Set_ComboList(oForm.Items.Item("CpCode").Specific, "SELECT '',''", "", true, false);
                                oForm.Items.Item("CpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP052_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP052_AddMatrixRow02(0, true);
                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                oMat04.LoadFromDataSource();
                                PS_PP052_AddMatrixRow04(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                oForm.Items.Item("ItemCode").Specific.Value = "";
                                oForm.Items.Item("ItemName").Specific.Value = "";
                                oForm.Items.Item("OrdNum").Specific.Value = "";
                                oForm.Items.Item("OrdSub1").Specific.Value = "";
                                oForm.Items.Item("OrdSub2").Specific.Value = "";
                                oForm.Items.Item("PP030HNo").Specific.Value = "";
                                dataHelpClass.Set_ComboList(oForm.Items.Item("CpCode").Specific, "SELECT '',''", "", true, false);
                                oForm.Items.Item("CpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP052_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP052_AddMatrixRow02(0, true);
                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                oMat04.LoadFromDataSource();
                                PS_PP052_AddMatrixRow04(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else
                            {
                                if (pVal.ItemUID != "CardType")
                                {
                                    oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                }
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oMat02.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat03.LoadFromDataSource();
                        oMat03.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oMat04.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Settings.MatrixUID = "Mat03";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oMat04.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Opt03")
                    {
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oMat04.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Opt04")
                    {
                        oForm.Settings.MatrixUID = "Mat04";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oMat04.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oMat01Row01 = pVal.Row;
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                            oMat02Row02 = pVal.Row;
                        }
                    }
                    if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat03.SelectRow(pVal.Row, true, false);
                            oMat03Row03 = pVal.Row;
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "LBtn01")
                    {
                        PS_PP030 PS_PP030 = new PS_PP030();
                        PS_PP030.LoadForm(oForm.Items.Item("PP030HNo").Specific.Value);
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            PS_PP030 PS_PP030 = new PS_PP030();
                            PS_PP030.LoadForm(codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.IndexOf("-") - 1));
                        }
                        if (pVal.ColUID == "PP030HNo")
                        {
                            PS_PP030 PS_PP030 = new PS_PP030();
                            PS_PP030.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            PS_PP030 PS_PP030 = new PS_PP030();
                            PS_PP030.LoadForm(codeHelpClass.Mid(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.IndexOf("-") - 1));
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
            int Hour_Renamed;
            int Minute_Renamed;
            double Weight;
            double Time;
            string Query01;
            string errMessage = string.Empty;
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
                            if (PS_PP052_Validate("수정01") == false)
                            {
                                oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP052L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrdMgNum")
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                    }
                                    else
                                    {
                                        if (oForm.Items.Item("PP030HNo").Specific.Value != codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim().IndexOf("-")))
                                        {
                                            oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                        }
                                        else
                                        {
                                            Query01 = "EXEC PS_PP040_01 '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "', '" + oForm.Items.Item("OrdType").Specific.Selected.Value + "'";
                                            oRecordSet01.DoQuery(Query01);
                                            if (oRecordSet01.RecordCount == 0)
                                            {
                                                oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else
                                            {
                                                oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oRecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP052L.SetValue("U_Sequence", pVal.Row - 1, oRecordSet01.Fields.Item("Sequence").Value);
                                                oDS_PS_PP052L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP052L.SetValue("U_CpName", pVal.Row - 1, oRecordSet01.Fields.Item("CpName").Value);
                                                oDS_PS_PP052L.SetValue("U_OrdGbn", pVal.Row - 1, oRecordSet01.Fields.Item("OrdGbn").Value);
                                                oDS_PS_PP052L.SetValue("U_BPLId", pVal.Row - 1, oRecordSet01.Fields.Item("BPLId").Value);
                                                oDS_PS_PP052L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP052L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                                oDS_PS_PP052L.SetValue("U_OrdNum", pVal.Row - 1, oRecordSet01.Fields.Item("OrdNum").Value);
                                                oDS_PS_PP052L.SetValue("U_OrdSub1", pVal.Row - 1, oRecordSet01.Fields.Item("OrdSub1").Value);
                                                oDS_PS_PP052L.SetValue("U_OrdSub2", pVal.Row - 1, oRecordSet01.Fields.Item("OrdSub2").Value);
                                                oDS_PS_PP052L.SetValue("U_PP030HNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030HNo").Value);
                                                oDS_PS_PP052L.SetValue("U_PP030MNo", pVal.Row - 1, oRecordSet01.Fields.Item("PP030MNo").Value);
                                                oDS_PS_PP052L.SetValue("U_SelWt", pVal.Row - 1, oRecordSet01.Fields.Item("SelWt").Value);
                                                oDS_PS_PP052L.SetValue("U_PSum", pVal.Row - 1, oRecordSet01.Fields.Item("PSum").Value);
                                                oDS_PS_PP052L.SetValue("U_BQty", pVal.Row - 1, oRecordSet01.Fields.Item("BQty").Value);
                                                oDS_PS_PP052L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                oDS_PS_PP052L.SetValue("U_LineId", pVal.Row - 1, "");
                                                oDS_PS_PP052L.SetValue("U_MachCode", pVal.Row - 1, "");
                                                oDS_PS_PP052L.SetValue("U_MachName", pVal.Row - 1, "");
                                                if (oMat03.VisualRowCount == 0)
                                                {
                                                    PS_PP052_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    PS_PP052_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                }

                                                oDS_PS_PP052N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oRecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP052N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oRecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP052N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oRecordSet01.Fields.Item("CpName").Value);
                                                oDS_PS_PP052N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
                                            }

                                            if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP052L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                            {
                                                PS_PP052_AddMatrixRow01(pVal.Row, false);
                                            }
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "PQty")
                                {
                                    double test = Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP052L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        oDS_PS_PP052L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        Weight = Convert.ToDouble(dataHelpClass.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
                                        if (Weight == 0)
                                        {
                                            oDS_PS_PP052L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP052L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else
                                        {
                                            oDS_PS_PP052L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP052L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        oDS_PS_PP052L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP052L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }
                                else if (pVal.ColUID == "NQty")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP052L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP052L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                    }
                                }
                                else if (pVal.ColUID == "WorkTime")
                                {
                                    if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1")
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToInt32(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else
                                    {
                                        oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToInt32(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                }
                                else if (pVal.ColUID == "MachCode")
                                {
                                    oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP052L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "CItemCod")
                                {
                                    oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP052L.SetValue("U_CItemNam", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP052L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "WorkCode")
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP052M.SetValue("U_WorkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                oDS_PS_PP052M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value.ToString().Trim());

                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP052M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP052_AddMatrixRow02(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "NStart")
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToInt32(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                if (Convert.ToInt32(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToInt32(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP052M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP052M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value)));
                                    oDS_PS_PP052M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value)));
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        Time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        Time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    Hour_Renamed = Convert.ToInt32(Time / 100);
                                    Minute_Renamed = Convert.ToInt32(Time % 100);
                                    Time = Hour_Renamed;
                                    if (Minute_Renamed > 0)
                                    {
                                        Time += 0.5;
                                    }
                                    oDS_PS_PP052M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
                                    oDS_PS_PP052M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                    oDS_PS_PP052M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                }
                            }
                            else if (pVal.ColUID == "NEnd")
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP052M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP052M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value)));
                                    oDS_PS_PP052M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value)));
                                }
                                else
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        Time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        Time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    Hour_Renamed = Convert.ToInt32(Time / 100);
                                    Minute_Renamed = Convert.ToInt32(Time % 100);
                                    Time = Hour_Renamed;
                                    if (Minute_Renamed > 0)
                                    {
                                        Time += 0.5;
                                    }
                                    oDS_PS_PP052M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
                                    oDS_PS_PP052M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                    oDS_PS_PP052M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - Time));
                                }
                            }
                            else if (pVal.ColUID == "YTime")
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                oDS_PS_PP052M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                            }
                            else
                            {
                                oDS_PS_PP052M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP052N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP052N.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP052N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP052H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BaseTime")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, Convert.ToString(Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value)));
                            }
                            else if (pVal.ItemUID == "OrdMgNum")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    PS_PP052_OrderInfoLoad();
                                }
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                oDS_PS_PP052H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP052_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP052_AddMatrixRow02(0, true);
                                oMat04.Clear();
                                oMat04.FlushToDataSource();
                                oMat04.LoadFromDataSource();
                                PS_PP052_AddMatrixRow04(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "UseMCode")
                            {
                                Query01 = "EXEC PS_PP040_98 '" + oForm.Items.Item("UseMCode").Specific.Value;
                                oRecordSet01.DoQuery(Query01);
                                oForm.Items.Item("UseMName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oMat02.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat03.LoadFromDataSource();
                        oMat03.AutoResizeColumns();
                        oMat04.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    PS_PP052_FormItemEnabled();
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_PP052_AddMatrixRow01(oMat01.VisualRowCount, false);
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        PS_PP052_AddMatrixRow02(oMat02.VisualRowCount, false);
                    }
                    else if (pVal.ItemUID == "Mat04")
                    {
                        PS_PP052_AddMatrixRow04(oMat04.VisualRowCount, false);
                    }
                }
                PS_PP052_SumWorkTime();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat04);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP052H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP052L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP052M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP052N);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP052W);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItemCode")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP040H", "U_ItemCode,U_ItemName", pVal.ItemUID, (short)pVal.Row, "", "", "");
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                        PS_PP052_AddMatrixRow01(0, true);
                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP052_AddMatrixRow02(0, true);
                        oMat04.Clear();
                        oMat04.FlushToDataSource();
                        oMat04.LoadFromDataSource();
                        PS_PP052_AddMatrixRow04(0, true);
                        oMat03.Clear();
                        oMat03.FlushToDataSource();
                        oMat03.LoadFromDataSource();
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP052_FormResize();
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
            int i = 0;
            int j;
            bool Exist;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (oForm.Mode != BoFormMode.fm_ADD_MODE)
                    {
                        errMessage = "입력후 수정할 수 없습니다.";
                        BubbleEvent = false;
                        throw new Exception();
                    }
                    if (oLastItemUID01 == "Mat01")
                    {
                        if (PS_PP052_Validate("행삭제01") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        for (i = 1; i <= oMat03.RowCount; i++)
                        {
                            if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value)
                            {
                                oDS_PS_PP052N.RemoveRecord(i - 1);
                                oMat03.DeleteRow(i);
                                oMat03.FlushToDataSource();
                                i = 1;
                                continue;
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (oLastItemUID01 == "Mat01")
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        for (i = 1; i <= oMat03.VisualRowCount; i++)
                        {
                            if (Convert.ToInt32(oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value) != 1)
                            {
                                oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value = Convert.ToString(Convert.ToInt32(oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value) - 1);
                            }
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_PP052L.RemoveRecord(oDS_PS_PP052L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP052_AddMatrixRow01(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP052L.GetValue("U_OrdMgNum", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP052_AddMatrixRow01(oMat01.RowCount, false);
                            }
                        }
                    }
                    else if (oLastItemUID01 == "Mat02")
                    {
                        for (i = 1; i <= oMat02.VisualRowCount; i++)
                        {
                            oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat02.FlushToDataSource();
                        oDS_PS_PP052M.RemoveRecord(oDS_PS_PP052M.Size - 1);
                        oMat02.LoadFromDataSource();
                        if (oMat02.RowCount == 0)
                        {
                            PS_PP052_AddMatrixRow02(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP052M.GetValue("U_WorkCode", oMat02.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP052_AddMatrixRow02(oMat02.RowCount, false);
                            }
                        }
                    }
                    else if (oLastItemUID01 == "Mat03")
                    {
                        for (i = 1; i <= oMat03.VisualRowCount; i++)
                        {
                            oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat03.FlushToDataSource();
                        //사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
                        if (oDS_PS_PP052N.Size == 1)
                        {
                        }
                        else
                        {
                            oDS_PS_PP052N.RemoveRecord(oDS_PS_PP052N.Size - 1);
                        }
                        oMat03.LoadFromDataSource();

                        //공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                        for (i = 1; i <= oMat01.RowCount - 1; i++)
                        {
                            Exist = false;
                            for (j = 1; j <= oMat03.RowCount; j++)
                            {
                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value && oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value)
                                {
                                    Exist = true;
                                }
                            }
                            //불량코드테이블에 값이 존재하지 않으면
                            if (Exist == false)
                            {
                                if (oMat03.VisualRowCount == 0)
                                {
                                    PS_PP052_AddMatrixRow03(0, true);
                                }
                                else
                                {
                                    PS_PP052_AddMatrixRow03(oMat03.VisualRowCount, false);
                                }
                                oDS_PS_PP052N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
                                oDS_PS_PP052N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
                                oDS_PS_PP052N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
                                oDS_PS_PP052N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(i));
                                oMat03.LoadFromDataSource();
                                oMat03.AutoResizeColumns();
                                oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                oMat03.FlushToDataSource();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
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
            finally
            {
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
                        Query01 += "   FROM [@PS_PP040H]";
                        Query01 += " WHERE U_DocType = '10'";
                        Query01 += " AND U_OrdGbn IN ('111','601')";
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
                        Query01 += "   FROM [@PS_PP040H]";
                        Query01 += " WHERE U_DocType = '10'";
                        Query01 += " AND U_OrdGbn IN ('111','601')";
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
                    Query01 += "  FROM [@PS_PP040H]";
                    Query01 += " WHERE U_DocType = '10'";
                    Query01 += "   AND U_OrdGbn IN ('111','601')";

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
                    Query01 += "  FROM [@PS_PP040H]";
                    Query01 += " WHERE U_DocType = '10'";
                    Query01 += "   AND U_OrdGbn IN ('111','601')";

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
                        case "1284": //취소
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_PP052_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80101" || oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80111")
                                {
                                    for (i = 1; i < oMat01.VisualRowCount; i++)
                                    {
                                        if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDoc").Cells.Item(i).Specific.Value))
                                        {
                                            PSH_Globals.SBO_Application.MessageBox("출고번호 미생성된 건은 취소불가! 관리자에게 문의하세요.");
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }

                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                if (oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80101" || oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() == "CP80111")
                                {
                                    if (PS_PP052_Add_InventoryGenEntry() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 취소할수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID,ref  pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
                            oMat01.AutoResizeColumns();
                            oMat02.AutoResizeColumns();
                            oMat04.AutoResizeColumns();
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
                            PS_PP052_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP052_FormItemEnabled();
                            PS_PP052_AddMatrixRow01(0, true);
                            PS_PP052_AddMatrixRow02(0, true);
                            PS_PP052_AddMatrixRow04(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
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
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (PS_PP052_FindValidateDocument("@PS_PP040H") == false)
                                {
                                    //찾기메뉴 활성화일때 수행
                                    if (PSH_Globals.SBO_Application.Menus.Item("1281").Enabled == true)
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                                    }
                                    else
                                    {
                                        PSH_Globals.SBO_Application.MessageBox("관리자에게 문의바랍니다");
                                    }
                                    BubbleEvent = false;
                                    return;
                                }
                            }
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
                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oMat01Row01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat02")
                {
                    if (pVal.Row > 0)
                    {
                        oMat02Row02 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat03")
                {
                    if (pVal.Row > 0)
                    {
                        oMat03Row03 = pVal.Row;
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
