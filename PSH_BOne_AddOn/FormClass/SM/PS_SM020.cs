using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 재고장(배치)
    /// </summary>
    internal class PS_SM020 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_SM020H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_SM020L; //등록라인
        private SAPbouiCOM.Form oBaseForm;
        private string oBaseItemUID;
        private string oBaseColUID;
        private int oBaseColRow;
        private string oBaseTradeType;
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oMat01Row01;
        private int oMat02Row02;

        /// <summary>
		/// Form 호출(Main Menu에서 호출)
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
        {
            LoadForm();
        }

        /// <summary>
        /// Form 호출(다른 폼에서 호출)
        /// </summary>
        /// <param name="baseForm">기준 Form</param>
        /// <param name="baseItemUID">기준 Form의 ItemUID</param>
        /// <param name="baseColUID">기준 Form의 Matrix ColUID</param>
        /// <param name="baseMatRow">기준 Form의 Matrix Row</param>
        /// <param name="baseTradeType">기준 Form의 TradeType</param>
        public void LoadForm(SAPbouiCOM.Form baseForm, string baseItemUID, string baseColUID, int baseMatRow, string baseTradeType)
        {
            oBaseForm = baseForm;
            oBaseItemUID = baseItemUID;
            oBaseColUID = baseColUID;
            oBaseColRow = baseMatRow;
            oBaseTradeType = baseTradeType;

            LoadForm();
        }

        /// <summary>
        /// Form 호출
        /// </summary>
        private void LoadForm()
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SM020.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PS_SM020_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SM020");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                PS_SM020_CreateItems();
                PS_SM020_SetComboBox();
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
        private void PS_SM020_CreateItems()
        {
            try
            {
                oDS_PS_SM020H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_SM020L = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                oForm.DataSources.UserDataSources.Add("StockType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("StockType").Specific.DataBind.SetBound(true, "", "StockType");

                oForm.DataSources.UserDataSources.Add("TradeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TradeType").Specific.DataBind.SetBound(true, "", "TradeType");

                oForm.DataSources.UserDataSources.Add("ItemGpCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItemGpCd").Specific.DataBind.SetBound(true, "", "ItemGpCd");

                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");

                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Mark").Specific.DataBind.SetBound(true, "", "Mark");

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");

                oForm.Items.Item("Mat01").Enabled = false;
                oForm.Items.Item("Mat02").Enabled = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SM020_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_SM020", "StockType", "", "1", "재고있는품목");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM020", "StockType", "", "2", "전체");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("StockType").Specific, "PS_SM020", "StockType", false);
                oForm.Items.Item("StockType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                dataHelpClass.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "", "전체");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM020", "TradeType", "", "2", "임가공");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("TradeType").Specific, "PS_SM020", "TradeType", false);
                oForm.Items.Item("TradeType").Specific.Select(oBaseTradeType, SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", false, false);

                //중분류는 대분류를 선택했을 때(COMBO_SELECT 이벤트) 재설정, 이 메서드에서는 구현 불필요
                //oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
                //dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", false, false);

                oForm.Items.Item("ItemType").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", false, false);

                oForm.Items.Item("Mark").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", false, false);

                oForm.Items.Item("ItemGpCd").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemGpCd").Specific, "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", false, false);

                oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Mark").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ItemGpCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("TradeType").Enabled = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Matrix 데이터 로드
        /// </summary>
        private void PS_SM020_MTX01()
        {
            string errMessage = string.Empty;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;
            string Param09;
            string Param10;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                oForm.Freeze(true);

                Param01 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("StockType").Specific.Selected.Value.ToString().Trim();
                Param03 = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
                Param04 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim();
                Param05 = oForm.Items.Item("ItmMsort").Specific.Selected.Value.ToString().Trim();
                Param06 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
                Param07 = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
                Param08 = oForm.Items.Item("Mark").Specific.Selected.Value.ToString().Trim();
                Param09 = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
                Param10 = oForm.Items.Item("ItemGpCd").Specific.Selected.Value.ToString().Trim();

                string baseFormType = oBaseForm != null ? oBaseForm.TypeEx : "";

                if
                (
                    baseFormType == "133" //AR송장
                    || baseFormType == "139" //판매오더
                    || baseFormType == "140" //납품
                    || baseFormType == "149" //판매견적
                    || baseFormType == "179" //AR대변메모
                    || baseFormType == "180" //반품(판매)
                    || baseFormType == "60091" //AR예약송장
                )
                {
                    Query01 = "EXEC PS_SM020_01 '" + Param01 + "','Y','','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
                }
                else if
                (
                    baseFormType == "141" //AP송장
                    || baseFormType == "142" //구매오더
                    || baseFormType == "143" //입고PO
                    || baseFormType == "181" //AP대변메모
                    || baseFormType == "182" //반품(구매)
                    || baseFormType == "60092" //AP예약송장
                )
                {
                    Query01 = "EXEC PS_SM020_01 '" + Param01 + "','','Y','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
                }
                else
                {
                    Query01 = "EXEC PS_SM020_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
                }

                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oForm.Items.Item("Mat01").Enabled = false;
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("Mat01").Enabled = true;
                }

                for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SM020H.InsertRecord(i);
                    }
                    oDS_PS_SM020H.Offset = i;
                    oDS_PS_SM020H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SM020H.SetValue("U_ColReg01", i, Convert.ToString(false));
                    oDS_PS_SM020H.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemCode").Value);
                    oDS_PS_SM020H.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("ItemName").Value);
                    oDS_PS_SM020H.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CallSize").Value);
                    oDS_PS_SM020H.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("Mark").Value);
                    oDS_PS_SM020H.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("OnHand").Value);
                    oDS_PS_SM020H.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("IsCommited").Value);
                    oDS_PS_SM020H.SetValue("U_ColQty03", i, RecordSet01.Fields.Item("OnOrder").Value);
                    oDS_PS_SM020H.SetValue("U_ColQty04", i, RecordSet01.Fields.Item("OnEnabled").Value);
                    oDS_PS_SM020H.SetValue("U_ColNum01", i, "0"); //선택수량

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// Matrix 데이터 로드
        /// </summary>
        private void PS_SM020_MTX02()
        {
            string errMessage = string.Empty;
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                oForm.Freeze(true);

                string Param01 = null;

                Param01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;

                Query01 = "EXEC PS_SM020_02 '" + Param01 + "'";
                RecordSet01.DoQuery(Query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oForm.Items.Item("Mat02").Enabled = false;
                    errMessage = "배치 결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("Mat02").Enabled = true;
                }

                if (dataHelpClass.GetItem_ItmBsort(Param01) == "104" || dataHelpClass.GetItem_ItmBsort(Param01) == "302")
                {
                    oMat02.Columns.Item("SelQty").Editable = false;
                    oMat02.Columns.Item("SelWeight").Editable = false;
                }
                else
                {
                    oMat02.Columns.Item("SelQty").Editable = true;
                    oMat02.Columns.Item("SelWeight").Editable = true;
                }

                for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SM020L.InsertRecord(i);
                    }
                    oDS_PS_SM020L.Offset = i;
                    oDS_PS_SM020L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_SM020L.SetValue("U_ColReg01", i, Convert.ToString(false));
                    oDS_PS_SM020L.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("BatchNum").Value);
                    oDS_PS_SM020L.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("WhsCode").Value);
                    oDS_PS_SM020L.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("WhsName").Value);
                    oDS_PS_SM020L.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("PackNo").Value);
                    oDS_PS_SM020L.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("Weight").Value);
                    oDS_PS_SM020L.SetValue("U_ColNum02", i, RecordSet01.Fields.Item("SelQty").Value);
                    oDS_PS_SM020L.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("SelWeight").Value);

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }

                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// BaseForm 설정
        /// </summary>
        private void PS_SM020_SetBaseForm()
        {
            string errMessage = string.Empty;
            string itemCode;
            SAPbouiCOM.Matrix oBaseMat = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oBaseForm != null)
                {
                    string baseFormType = oBaseForm.TypeEx;

                    if
                    (
                        baseFormType == "133" //AR송장
                        || baseFormType == "139" //판매오더
                        || baseFormType == "140" //납품
                        || baseFormType == "141" //AP송장
                        || baseFormType == "142" //구매오더
                        || baseFormType == "143" //입고PO
                        || baseFormType == "149" //판매견적
                        || baseFormType == "179" //AR대변메모
                        || baseFormType == "180" //반품(판매)
                        || baseFormType == "181" //AP대변메모
                        || baseFormType == "182" //반품(구매)
                        || baseFormType == "60091" //AR예약송장
                        || baseFormType == "60092" //AP예약송장
                    )
                    {
                        oBaseMat = oBaseForm.Items.Item("38").Specific;

                        for (int i = 1; i <= oMat01.RowCount; i++) //품목
                        {
                            if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.Value) <= 0)
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = itemCode;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oBaseColRow += 1;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = itemCode;
                                    oBaseMat.Columns.Item("U_Qty").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value); //수량
                                    oBaseMat.Columns.Item("U_Unweight").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)); //단중
                                    oBaseMat.Columns.Item("11").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.Value); //중량
                                    oBaseMat.Columns.Item("14").Cells.Item(oBaseColRow).Specific.Value = dataHelpClass.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm.Items.Item("4").Specific.Value + "','" + itemCode + "'", 0, 1);
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oBaseColRow += 1;
                                }
                            }
                        }

                        for (int i = 1; i <= oMat02.RowCount; i++) //배치
                        {
                            if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                            {
                                if (Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value) <= 0)
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = itemCode;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oBaseColRow += 1;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = itemCode; //품목
                                    oBaseMat.Columns.Item("U_Qty").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value); //수량
                                    oBaseMat.Columns.Item("U_Unweight").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)); //단중
                                    oBaseMat.Columns.Item("11").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value); //중량
                                    oBaseMat.Columns.Item("14").Cells.Item(oBaseColRow).Specific.Value = dataHelpClass.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm.Items.Item("4").Specific.Value + "','" + itemCode + "'", 0, 1);
                                    oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oBaseColRow += 1;
                                }
                            }
                        }
                    }
                    else if (baseFormType == "720") //출고
                    {
                        oBaseMat = oBaseForm.Items.Item("13").Specific; //출고-매트릭스

                        for (int i = 1; i <= oMat01.RowCount; i++) //품목매트릭스
                        {
                            if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                            {
                                oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value; //품목코드
                                oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oBaseColRow += 1;
                            }
                        }

                        for (int i = 1; i <= oMat02.RowCount; i++) //배치매트릭스
                        {
                            if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                            {
                                oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value; //품목코드
                                oBaseMat.Columns.Item("1").Cells.Item(oBaseColRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oBaseColRow += 1;
                            }
                        }
                    }
                    else if (baseFormType == "PS_SD091") //이동요청등록
                    {
                        oBaseMat = oBaseForm.Items.Item("Mat01").Specific; //매트릭스

                        for (int i = 1; i <= oMat01.RowCount; i++)
                        {
                            if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value) > 0)
                                {
                                    oBaseMat.Columns.Item("ItemCode").Cells.Item(oBaseColRow).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value; //품목
                                    oBaseMat.Columns.Item("Qty").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value); //수량
                                    oBaseMat.Columns.Item("Unweight").Cells.Item(oBaseColRow).Specific.Value = Convert.ToDouble(dataHelpClass.GetItem_UnWeight(oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value)); //단중
                                    oBaseMat.Columns.Item("Weight").Cells.Item(oBaseColRow).Specific.Value = System.Math.Round(Convert.ToDouble(dataHelpClass.GetItem_UnWeight(oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value)) * Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value) / 1000, 3); //단중
                                    oBaseColRow += 1;
                                }
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
                if (oBaseMat != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseMat);
                }
            }
        }

        /// <summary>
        /// 필수 데이터 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SM020_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                string baseFormType = oBaseForm.TypeEx;

                if (baseFormType == "720") //출고
                {
                    for (int i = 1; i <= oMat01.RowCount; i++)
                    {
                        if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                        {
                            if (Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value) <= 0)
                            {
                                errMessage = "품목 수량은 필수입니다.";
                                throw new Exception();
                            }
                        }
                    }

                    for (int i = 1; i <= oMat02.RowCount; i++)
                    {
                        if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                        {
                            if (Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value) <= 0)
                            {
                                errMessage = "배치 수량은 필수입니다.";
                                throw new Exception();
                            }
                        }
                    }
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_SM020_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_SM020_SetBaseForm(); //부모폼에입력
                            oForm.Close();
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemName", "");
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
                }
                else if (pVal.Before_Action == false)
                {
                }

                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItmBsort")
                    {
                        //기존 콤보 데이터 삭제
                        for (int i = oForm.Items.Item("ItmMsort").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("ItmMsort").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }

                        oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
                        dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("ItmBsort").Specific.Selected.Value + "' ORDER BY U_Code", "", false, false);

                        if (oForm.Items.Item("ItmMsort").Specific.ValidValues.Count > 0)
                        {
                            oForm.Items.Item("ItmMsort").Specific.Select("선택", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row > 0)
                            {
                                oMat01.SelectRow(pVal.Row, true, false);
                                oMat01Row01 = pVal.Row;

                                if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value) == "Y") //배치 사용 품목
                                {
                                    PS_SM020_MTX02();
                                }
                                else
                                {
                                    PS_SM020_MTX02();
                                }
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row > 0)
                            {
                                oMat02.SelectRow(pVal.Row, true, false);
                                oMat02Row02 = pVal.Row;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.Enabled = true;
                        oForm.Settings.EnableRowFormat = true;
                    }
                    else if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.Enabled = true;
                        oForm.Settings.EnableRowFormat = true;
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat02.FlushToDataSource();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string itemCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SelQty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
                                    oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value;

                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                                oForm.Update();
                            }
                            else if (pVal.ColUID == "SelWeight")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
                                    oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value;

                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                                oForm.Update();
                            }
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "SelQty")
                            {
                                if (Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
                                    oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;

                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                                oForm.Update();
                            }
                            else if (pVal.ColUID == "SelWeight")
                            {
                                if (Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
                                    oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
                                }
                                else
                                {
                                    itemCode = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;

                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                                oForm.Update();
                            }
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
                    oMat01.AutoResizeColumns();
                    oMat02.AutoResizeColumns();
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
                    if (oForm != null)
                    {
                        oForm = null;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SM020H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SM020L);
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oForm.Items.Item("Mat01").Top = 70;
                    oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 70;
                    oForm.Items.Item("Mat01").Left = 7;
                    oForm.Items.Item("Mat01").Width = oForm.Width - 21;
                    oForm.Items.Item("Mat02").Top = (oForm.Height / 2) + 10;
                    oForm.Items.Item("Mat02").Height = (oForm.Height / 2) - 75;
                    oForm.Items.Item("Mat02").Left = 7;
                    oForm.Items.Item("Mat02").Width = oForm.Width - 21;
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
                    case "Mat02":
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
