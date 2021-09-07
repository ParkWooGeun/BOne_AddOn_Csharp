using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산의뢰접수
	/// </summary>
	internal class PS_PP010 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_PP010H;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private string oLastRightClickDocEntry;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP010.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP010_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP010");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_PP010_CreateItems();
				PS_PP010_SetComboBox();
				//PS_PP010_AddMatrixRow(0, true);
				//PS_PP010_LoadCaption();
				//PS_PP010_EnableFormItem();

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", true); //취소
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP010_CreateItems()
        {
            try
            {
                oDS_PS_PP010H = oForm.DataSources.DBDataSources.Item("@PS_PP010H");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //품목대분류
                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                //작업구분
                oForm.DataSources.UserDataSources.Add("WorkGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
                oForm.Items.Item("WorkGbn").Specific.DataBind.SetBound(true, "", "WorkGbn");

                //고객
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                //고객명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

                //품목코드
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //생산요청일자(시작)
                oForm.DataSources.UserDataSources.Add("RegDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("RegDateFr").Specific.DataBind.SetBound(true, "", "RegDateFr");

                //생산요청일자(종료)
                oForm.DataSources.UserDataSources.Add("RegDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("RegDateTo").Specific.DataBind.SetBound(true, "", "RegDateTo");

                //생산의뢰일자(시작)
                oForm.DataSources.UserDataSources.Add("PuDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("PuDateFr").Specific.DataBind.SetBound(true, "", "PuDateFr");

                //생산의뢰일자(종료)
                oForm.DataSources.UserDataSources.Add("PuDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("PuDateTo").Specific.DataBind.SetBound(true, "", "PuDateTo");

                //작명
                oForm.DataSources.UserDataSources.Add("JakName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("JakName").Specific.DataBind.SetBound(true, "", "JakName");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP010_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.DataSources.UserDataSources.Item("BPLId").ValueEx = dataHelpClass.User_BPLID(); //COMBO_SELECT 이벤트 미발생

                //품목대분류
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("ItmBSort").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작업구분(헤더)
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("10", "영업");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("20", "정비");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("30", "멀티");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("40", "소재");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("50", "R/D");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("60", "견본");
                oForm.Items.Item("WorkGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작업구분(라인)
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("10", "영업");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("20", "정비");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("30", "멀티");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("40", "소재");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("50", "R/D");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("60", "견본");

                //사용처(라인)
                sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("UseDept").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
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
        /// 행 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_PP010_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_PP010H.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP010H.Offset = oRow;
                oDS_PS_PP010H.SetValue("DocNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 버튼 캡션 설정
        /// </summary>
        private void PS_PP010_LoadCaption()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "추가";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "확인";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "갱신";
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_PP010_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("WorkGbn").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("RegDateFr").Enabled = true;
                    oForm.Items.Item("RegDateTo").Enabled = true;
                    oForm.Items.Item("PuDateFr").Enabled = false;
                    oForm.Items.Item("PuDateTo").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;

                    oMat01.Columns.Item("RegNum").Editable = true;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("WorkGbn").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("RegDateFr").Enabled = false;
                    oForm.Items.Item("RegDateTo").Enabled = false;
                    oForm.Items.Item("PuDateFr").Enabled = true;
                    oForm.Items.Item("PuDateTo").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;

                    oMat01.Columns.Item("RegNum").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP010_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            string RegNum;
            string ItmBsort;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "ShipCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("ShipCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ShipName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "ItemCode":
                        sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "Mat01":
                        if (oCol == "RegNum")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("RegNum").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_PP010_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            RegNum = oMat01.Columns.Item("RegNum").Cells.Item(oRow).Specific.Value.ToString().Trim();
                            ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();

                            sQry = "  Select  IsNull(U_SjDocLin, '') ";
                            sQry += " From    [@PS_SD010H] ";
                            sQry += " Where   U_RegNum = '" + RegNum + "'";
                            sQry += "         AND U_ItmBSort = '" + ItmBsort + "'";

                            oRecordSet01.DoQuery(sQry);

                            if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                            {
                                sQry = "  Select  a.U_ItemCode, ";
                                sQry += "         a.U_ItemName, ";
                                sQry += "         b.U_Material, ";
                                sQry += "         b.SalUnitMsr, ";
                                sQry += "         b.U_Size, ";
                                sQry += "         a.U_ItmBSort, ";
                                sQry += "         e.U_CPNaming, ";
                                sQry += "         a.U_ReWeight ";
                                sQry += " From    [@PS_SD010H] a";
                                sQry += "         Inner Join ";
                                sQry += "         [OITM] b ";
                                sQry += "             On a.U_ItemCode = b.ItemCode ";
                                sQry += "         Left Join ";
                                sQry += "         [@PSH_ItmMSort] e ";
                                sQry += "             On b.U_ItmBSort = e.U_rCode ";
                                sQry += "             And b.U_ItmMsort = e.U_Code ";
                                sQry += " Where   a.U_RegNum = '" + RegNum + "'";
                                sQry += "         AND a.U_ItmBSort = '" + ItmBsort + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                            else
                            {
                                sQry = "  Select  a.U_ItemCode, ";
                                sQry += "         a.U_ItemName, ";
                                sQry += "         b.U_Material, ";
                                sQry += "         b.SalUnitMsr, ";
                                sQry += "         b.U_Size, ";
                                sQry += "         a.U_ItmBSort, ";
                                sQry += "         e.U_CPNaming, ";
                                sQry += "         a.U_ReWeight, ";
                                sQry += "         Left(a.U_SjDocLin, CharIndex('-', a.U_SjDocLin) - 1) As SjDocNum, ";
                                sQry += "         Right(a.U_SjDocLin, Len(a.U_SjDocLin) - CharIndex('-', a.U_SjDocLin)) As SjLinNum, ";
                                sQry += "         a.U_SjQty, ";
                                sQry += "         a.U_SjWeight, ";
                                sQry += "         c.DocDate, ";
                                sQry += "         d.ShipDate, ";
                                sQry += "         d.LineTotal, ";
                                sQry += "         a.U_CardCode, ";
                                sQry += "         a.U_CardName ";
                                sQry += " From    [@PS_SD010H] a ";
                                sQry += "         Inner Join ";
                                sQry += "         [OITM] b ";
                                sQry += "             On a.U_ItemCode = b.ItemCode ";
                                sQry += "         Inner Join ";
                                sQry += "         [ORDR] c ";
                                sQry += "             On c.DocNum = Left(a.U_SjDocLin, CharIndex('-', a.U_SjDocLin) - 1) ";
                                sQry += "         Inner Join ";
                                sQry += "         [RDR1] d ";
                                sQry += "             On c.DocEntry = d.DocEntry ";
                                sQry += "             And d.LineNum = Right(a.U_SjDocLin, Len(a.U_SjDocLin) - CharIndex('-', a.U_SjDocLin)) ";
                                sQry += "         Left Join ";
                                sQry += "         [@PSH_ItmMSort] e ";
                                sQry += "             On b.U_ItmBSort = e.U_rCode ";
                                sQry += "             And b.U_ItmMsort = e.U_Code ";
                                sQry += " Where   a.U_RegNum = '" + RegNum + "'";
                                sQry += "         AND a.U_ItmBSort = '" + ItmBsort + "'";

                                oRecordSet01.DoQuery(sQry);

                                oMat01.Columns.Item("SjDocNum").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("SjDocNum").Value.ToString().Trim();
                                oMat01.Columns.Item("SjLinNum").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("SjLinNum").Value.ToString().Trim();
                                oMat01.Columns.Item("SjQty").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjQty").Value.ToString().Trim();
                                oMat01.Columns.Item("SjWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjWeight").Value.ToString().Trim();
                                oMat01.Columns.Item("SjDcDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                                oMat01.Columns.Item("SjDuDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ShipDate").Value.ToString("yyyyMMdd");
                                oMat01.Columns.Item("SlePrice").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim();
                                oMat01.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim();
                                oMat01.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim();
                            }

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim();
                            oMat01.Columns.Item("Material").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Material").Value.ToString().Trim();
                            oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("SalUnitMsr").Value.ToString().Trim();
                            oMat01.Columns.Item("Size").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim();
                            oMat01.Columns.Item("ItmBSort").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("CpName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CpNaming").Value.ToString().Trim();
                            oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("PuDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("ProDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd"); 
                            oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("SubNo1").Cells.Item(oRow).Specific.Value = "00";
                            oMat01.Columns.Item("SubNo2").Cells.Item(oRow).Specific.Value = "000";
                            oMat01.Columns.Item("ReWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ReWeight").Value.ToString().Trim();

                            oMat01.Columns.Item("Comments").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "ItemCode")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_PP010_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            sQry = "  Select  a.ItemCode, ";
                            sQry += "         a.FrgnName, ";
                            sQry += "         a.U_Material, ";
                            sQry += "         a.SalUnitMsr, ";
                            sQry += "         a.U_Size, ";
                            sQry += "         a.U_ItmBSort, ";
                            sQry += "         b.U_CPNaming  ";
                            sQry += " From    OITM a ";
                            sQry += "         Left Join ";
                            sQry += "         [@PSH_ItmMSort] b ";
                            sQry += "             On a.U_ItmBSort = b.U_rCode ";
                            sQry += "             And a.U_ItmMsort = b.U_Code ";
                            sQry += " Where   a.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("FrgnName").Value.ToString().Trim(); //품명
                            oMat01.Columns.Item("Material").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Material").Value.ToString().Trim(); //재질
                            oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("SalUnitMsr").Value.ToString().Trim(); //단위
                            oMat01.Columns.Item("Size").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim(); //규격
                            oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim(); //작번

                            oMat01.Columns.Item("PuDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("ItmBSort").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "ShipCode")
                        {
                            sQry = "  Select  CardName ";
                            sQry += " From    OCRD ";
                            sQry += " Where   CardCode = '" + oMat01.Columns.Item("ShipCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("ShipName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        oMat01.AutoResizeColumns();
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Header 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP010_CheckHeaderDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            
            return returnValue;
        }

        /// <summary>
        /// Line 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP010_CheckLineDataValid()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                
                if (oForm.Items.Item("WorkGbn").Specific.Selected.Value == "10") //구분이 영업(10)일 때만 요청번호 체크
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        if (string.IsNullOrEmpty(oDS_PS_PP010H.GetValue("U_RegNum", i)))
                        {
                            errMessage = (i + 1) + "번 라인의 생산요청번호가 없습니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP010_AddData()
        {
            bool returnValue = false;
            short loopCount;
            string ErrOrdNum = string.Empty; //선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
            string Query01;
            string Query02;
            string DocEntry;
            string BPLId;
            string RegNum;
            string ItemCode;
            string ItemName;
            string Material;
            string Unit;
            string Size;
            string ItmBsort;
            string CpName;
            string SjDocNum;
            string SjLinNum;
            string SjQty;
            string SjWeight;
            string SjDcDate;
            string SjDuDate;
            string SlePrice;
            string WorkGbn;
            string CardCode;
            string CardName;
            string ShipCode;
            string ShipName;
            string PuDate;
            string ProDate;
            string Comments;
            string JakName;
            string SubNo1;
            string SubNo2;
            string UseDept;
            string ReWeight;
            string Status;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            short MinusNum = 0;
            string errMessage = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    MinusNum = 2;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    MinusNum = 1;
                }

                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - MinusNum; loopCount++)
                {
                    DocEntry = oDS_PS_PP010H.GetValue("DocEntry", loopCount).ToString().Trim(); //DocEntry는 프로시저에서 처리
                    RegNum = oDS_PS_PP010H.GetValue("U_RegNum", loopCount).ToString().Trim();
                    ItemCode = oDS_PS_PP010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    ItemName = dataHelpClass.Make_ItemName(oDS_PS_PP010H.GetValue("U_ItemName", loopCount).ToString().Trim());
                    Material = oDS_PS_PP010H.GetValue("U_Material", loopCount).ToString().Trim();
                    Unit = oDS_PS_PP010H.GetValue("U_Unit", loopCount).ToString().Trim();
                    Size = oDS_PS_PP010H.GetValue("U_Size", loopCount).ToString().Trim();
                    ItmBsort = oDS_PS_PP010H.GetValue("U_ItmBSort", loopCount).ToString().Trim();
                    CpName = oDS_PS_PP010H.GetValue("U_CpName", loopCount).ToString().Trim();
                    SjDocNum = oDS_PS_PP010H.GetValue("U_SjDocNum", loopCount).ToString().Trim();
                    SjLinNum = oDS_PS_PP010H.GetValue("U_SjLinNum", loopCount).ToString().Trim();
                    SjQty = string.IsNullOrEmpty(oDS_PS_PP010H.GetValue("U_SjQty", loopCount).ToString().Trim()) ? "0" : oDS_PS_PP010H.GetValue("U_SjQty", loopCount).ToString().Trim();
                    SjWeight = oDS_PS_PP010H.GetValue("U_SjWeight", loopCount).ToString().Trim();
                    SjDcDate = oDS_PS_PP010H.GetValue("U_SjDcDate", loopCount).ToString().Trim();
                    SjDuDate = oDS_PS_PP010H.GetValue("U_SjDuDate", loopCount).ToString().Trim();
                    SlePrice = oDS_PS_PP010H.GetValue("U_SlePrice", loopCount).ToString().Trim();
                    CardCode = oDS_PS_PP010H.GetValue("U_CardCode", loopCount).ToString().Trim();
                    CardName = oDS_PS_PP010H.GetValue("U_CardName", loopCount).ToString().Trim();
                    ShipCode = oDS_PS_PP010H.GetValue("U_ShipCode", loopCount).ToString().Trim();
                    ShipName = oDS_PS_PP010H.GetValue("U_ShipName", loopCount).ToString().Trim();
                    PuDate = oDS_PS_PP010H.GetValue("U_PuDate", loopCount).ToString().Trim();
                    ProDate = oDS_PS_PP010H.GetValue("U_ProDate", loopCount).ToString().Trim();
                    Comments = oDS_PS_PP010H.GetValue("U_Comments", loopCount).ToString().Trim();
                    SubNo1 = string.IsNullOrEmpty(oDS_PS_PP010H.GetValue("U_SubNo1", loopCount).ToString().Trim()) ? "00" : oDS_PS_PP010H.GetValue("U_SubNo1", loopCount).ToString().Trim();
                    SubNo2 = string.IsNullOrEmpty(oDS_PS_PP010H.GetValue("U_SubNo2", loopCount).ToString().Trim()) ? "000" : oDS_PS_PP010H.GetValue("U_SubNo2", loopCount).ToString().Trim();
                    UseDept = oDS_PS_PP010H.GetValue("U_UseDept", loopCount).ToString().Trim();
                    ReWeight = oDS_PS_PP010H.GetValue("U_ReWeight", loopCount).ToString().Trim();

                    if (codeHelpClass.Left(ItemCode, 1) == "R")
                    {
                        WorkGbn = "50";
                        JakName = oDS_PS_PP010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    }
                    else if (codeHelpClass.Left(ItemCode, 1) == "E")
                    {
                        WorkGbn = "40";
                        JakName = oDS_PS_PP010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    }
                    else if (codeHelpClass.Left(ItemCode, 1) == "F")
                    {
                        WorkGbn = "30";
                        JakName = oDS_PS_PP010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    }
                    else if (codeHelpClass.Left(ItemCode, 1) == "X")
                    {
                        WorkGbn = "20";
                        JakName = oDS_PS_PP010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    }
                    else
                    {
                        WorkGbn = "10";
                        JakName = oDS_PS_PP010H.GetValue("U_JakName", loopCount).ToString().Trim();
                    }
                    Status = "O";

                    Query01 = "EXEC PS_PP010_03 '";
                    Query01 += DocEntry + "','"; //문서번호
                    Query01 += BPLId + "','"; //사업장코드
                    Query01 += RegNum + "','"; //생산요청번호
                    Query01 += ItemCode + "','"; //품목코드
                    Query01 += ItemName + "','"; //품목명
                    Query01 += Material + "','"; //재질
                    Query01 += Unit + "','"; //단위
                    Query01 += Size + "','"; //규격
                    Query01 += ItmBsort + "','"; //품목대분류
                    Query01 += CpName + "','"; //생성코드
                    Query01 += SjDocNum + "','"; //수주번호
                    Query01 += SjLinNum + "','"; //수주행
                    Query01 += SjQty + "','"; //수주량
                    Query01 += SjWeight + "','"; //수주중량
                    Query01 += SjDcDate + "','"; //수주일
                    Query01 += SjDuDate + "','"; //납기일
                    Query01 += SlePrice + "','"; //영업금액
                    Query01 += WorkGbn + "','";  //작업구분
                    Query01 += CardCode + "','"; //거래처코드
                    Query01 += CardName + "','"; //거래처명
                    Query01 += ShipCode + "','"; //납품처코드
                    Query01 += ShipName + "','"; //납품처명
                    Query01 += PuDate + "','"; //생산의뢰접수일
                    Query01 += ProDate + "','"; //작업지시일
                    Query01 += Comments + "','"; //비고
                    Query01 += JakName + "','"; //작명
                    Query01 += SubNo1 + "','"; //서브작번1
                    Query01 += SubNo2 + "','"; //서브작번2
                    Query01 += UseDept + "','"; //사용처
                    Query01 += ReWeight + "','"; //생산요청수량
                    Query01 += Status + "'"; //상태

                    //선행프로세스 대비 일자체크_S
                    BaseEntry = RegNum;
                    BaseLine = "0";
                    DocType = "PS_PP010";
                    CurDocDate = PuDate;

                    Query02 = "EXEC PS_Z_CHECK_DATE '";
                    Query02 += BaseEntry + "','";
                    Query02 += BaseLine + "','";
                    Query02 += DocType + "','";
                    Query02 += CurDocDate + "'";

                    RecordSet02.DoQuery(Query02);
                    //선행프로세스 대비 일자체크_E

                    if (RecordSet02.Fields.Item("ReturnValue").Value == "True")
                    {
                        RecordSet01.DoQuery(Query01); //등록
                    }
                    else
                    {
                        ErrOrdNum += " [" + ItemCode + "]";
                    }
                }
                
                if (!string.IsNullOrEmpty(ErrOrdNum)) //하나라도 선행프로세스 일자가 빠른 작번이 있으면
                {
                    errMessage = "생산의뢰접수일은 생산요청일과 같거나 늦어야합니다. 확인하십시오." + (char)13 + ErrOrdNum;
                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if(returnValue == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_PP010_LoadData()
        {
            short i;
            string sQry;
            string ItmBsort;
            string ItemCode;
            string CardCode;
            string WorkGbn;
            string BPLId;
            string PuDateFr;
            string PuDateTo;
            string JakName;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                PuDateFr = oForm.Items.Item("PuDateFr").Specific.Value.ToString().Trim();
                PuDateTo = oForm.Items.Item("PuDateTo").Specific.Value.ToString().Trim();
                JakName = oForm.Items.Item("JakName").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(BPLId))
                {
                    BPLId = "%";
                }

                if (string.IsNullOrEmpty(CardCode))
                {
                    CardCode = "%";
                }

                if (string.IsNullOrEmpty(ItemCode))
                {
                    ItemCode = "%";
                }

                if (string.IsNullOrEmpty(PuDateFr))
                {
                    PuDateFr = "19000101";
                }

                if (string.IsNullOrEmpty(PuDateTo))
                {
                    PuDateTo = "20991231";
                }

                if (string.IsNullOrEmpty(JakName))
                {
                    JakName = "%";
                }

                sQry = "EXEC [PS_PP010_01] '";
                sQry += BPLId + "','";
                sQry += ItmBsort + "','";
                sQry += WorkGbn + "','";
                sQry += CardCode + "','";
                sQry += ItemCode + "','";
                sQry += PuDateFr + "','";
                sQry += PuDateTo + "','";
                sQry += JakName + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_PP010H.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP010H.Size)
                    {
                        oDS_PS_PP010H.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_PP010H.Offset = i;
                    oDS_PS_PP010H.SetValue("DocNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP010H.SetValue("U_RegNum", i, oRecordSet01.Fields.Item("U_RegNum").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_Material", i, oRecordSet01.Fields.Item("U_Material").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_Unit", i, oRecordSet01.Fields.Item("U_Unit").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_Size", i, oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ItmBSort", i, oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_CpName", i, oRecordSet01.Fields.Item("U_CpName").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SjDocNum", i, oRecordSet01.Fields.Item("U_SjDocNum").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SjLinNum", i, oRecordSet01.Fields.Item("U_SjLinNum").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SjQty", i, oRecordSet01.Fields.Item("U_SjQty").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SjWeight", i, oRecordSet01.Fields.Item("U_SjWeight").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SjDcDate", i, oRecordSet01.Fields.Item("U_SjDcDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP010H.SetValue("U_SjDuDate", i, oRecordSet01.Fields.Item("U_SjDuDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP010H.SetValue("U_SlePrice", i, oRecordSet01.Fields.Item("U_SlePrice").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_WorkGbn", i, oRecordSet01.Fields.Item("U_WorkGbn").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_CardName", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ShipCode", i, oRecordSet01.Fields.Item("U_ShipCode").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ShipName", i, oRecordSet01.Fields.Item("U_ShipName").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_PuDate", i, oRecordSet01.Fields.Item("U_PuDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP010H.SetValue("U_ProDate", i, oRecordSet01.Fields.Item("U_ProDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP010H.SetValue("U_Comments", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_JakName", i, oRecordSet01.Fields.Item("U_JakName").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SubNo1", i, oRecordSet01.Fields.Item("U_SubNo1").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_SubNo2", i, oRecordSet01.Fields.Item("U_SubNo2").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_UseDept", i, oRecordSet01.Fields.Item("U_UseDept").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_ReWeight", i, oRecordSet01.Fields.Item("U_ReWeight").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("U_Status", i, oRecordSet01.Fields.Item("U_Status").Value.ToString().Trim());
                    oDS_PS_PP010H.SetValue("DocEntry", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
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
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP010_CheckHeaderDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_PP010_CheckLineDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP010_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Freeze(true);
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_PP010_AddMatrixRow(0, true);
                            oForm.Freeze(false);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP010_CheckLineDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP010_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_PP010_LoadCaption();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (PS_PP010_CheckHeaderDataValid() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PS_PP010_LoadCaption();
                        PS_PP010_LoadData();
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                            {
                                PS_SM010 tempForm = new PS_SM010();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 tempForm = new PS_SM010();
                                    tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "RegNum")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("RegNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "ShipCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ShipCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
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
            int sCount;
            int sSeq;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "BPLId")
                    {
                        if (oForm.Items.Item("BPLId").Specific.Value == "2" || oForm.Items.Item("BPLId").Specific.Value == "3")
                        {
                            sCount = oForm.Items.Item("ItmBSort").Specific.ValidValues.Count;
                            sSeq = sCount;
                            for (int i = 1; i <= sCount; i++)
                            {
                                oForm.Items.Item("ItmBSort").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                sSeq -= 1;
                            }
                            
                            if (oForm.Items.Item("BPLId").Specific.Selected.Value == "2")
                            {
                                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code";
                            }
                            else
                            {
                                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('108', '109') Order by Code";
                            }

                            oRecordSet01.DoQuery(sQry);
                            while (!oRecordSet01.EoF)
                            {
                                oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                                oRecordSet01.MoveNext();
                            }
                            oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_PP010_LoadCaption();
                        }
                    }

                    if (pVal.ItemUID == "WorkGbn" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();

                        PS_PP010_AddMatrixRow(0, true);
                        if (oForm.Items.Item("WorkGbn").Specific.Selected.Value == "10")
                        {
                            oMat01.Columns.Item("RegNum").Editable = true;
                            oMat01.Columns.Item("ItemCode").Editable = false;
                        }
                        else
                        {
                            oMat01.Columns.Item("RegNum").Editable = false;
                            oMat01.Columns.Item("ItemCode").Editable = true;
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_PP010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CardCode")
                        {
                            PS_PP010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            PS_PP010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "RegNum")
                            {
                                PS_PP010_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                PS_PP010_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ShipCode")
                            {
                                PS_PP010_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                            }
                            else
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                PS_PP010_LoadCaption();
                            }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP010H);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                            if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 생산의뢰접수를 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "Yes", "No") == 1)
                            {
                                //작번등록여부 체크
                                sQry = "  SELECT  COUNT(*) AS [Cnt]";
                                sQry += " FROM    [@PS_PP020H]";
                                sQry += " WHERE   U_PP010Doc = '" + oLastRightClickDocEntry + "'";
                                sQry += "         AND [Status] = 'O'";
                                oRecordSet01.DoQuery(sQry);

                                //작번등록이 존재하면
                                if (Convert.ToInt16(oRecordSet01.Fields.Item("Cnt").Value) > 0)
                                {
                                    PSH_Globals.SBO_Application.MessageBox("작번등록이 존재하는 작번입니다. 삭제할 수 없습니다.");
                                    BubbleEvent = false;
                                    return;
                                }
                                
                                sQry = "SELECT U_RegNum AS [RegNum] FROM [@PS_PP010H] WHERE DocEntry = '" + oLastRightClickDocEntry + "'";
                                oRecordSet01.DoQuery(sQry); //생산요청 번호 조회

                                sQry = "UPDATE [@PS_SD010H] SET U_Status = 'O' WHERE DocEntry = '" + oRecordSet01.Fields.Item("RegNum").Value + "'"; 
                                oRecordSet01.DoQuery(sQry); //생산요청[PS_SD010H]의 U_Status를 "O"로 복원

                                sQry = "Delete [@PS_PP010H] Where DocEntry = '" + oLastRightClickDocEntry + "'";
                                oRecordSet01.DoQuery(sQry); //생산의뢰접수 테이블에서 삭제

                                PSH_Globals.SBO_Application.StatusBar.SetText("삭제되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                                oLastRightClickDocEntry = "0";
                            }
                            else 
                            {
                                BubbleEvent = false;
                            }
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
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_PP010H.RemoveRecord(oDS_PS_PP010H.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                                oForm.Freeze(false);
                            }
                            break;
                        case "1281": //찾기
                            PS_PP010_EnableFormItem();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_PP010_LoadCaption();
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_PP010_EnableFormItem();

                            if (oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim() == "10")
                            {
                                oMat01.Columns.Item("RegNum").Editable = true;
                                oMat01.Columns.Item("ItemCode").Editable = false;
                            }
                            else
                            {
                                oMat01.Columns.Item("RegNum").Editable = false;
                                oMat01.Columns.Item("ItemCode").Editable = true;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_PP010_AddMatrixRow(0, true);
                            PS_PP010_LoadCaption();
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP010_EnableFormItem();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_PP010H.GetValue("Status", 0) == "O")
                                    {
                                        PS_PP010_AddMatrixRow(oMat01.RowCount, false);
                                    }
                                }
                            }
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
                    if (pVal.Row > 0)
                    {
                        oLastRightClickDocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
                    }
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
