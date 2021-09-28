using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기성매출등록
	/// </summary>
	internal class PS_SD050 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_SD050H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD050L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		
		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD050.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD050_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD050");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD050_CreateItems();
				PS_SD050_SetComboBox();
				PS_SD050_AddMatrixRow(0, true);
				PS_SD050_LoadCaption();
				PS_SD050_EnableFormItem();
				PS_SD050_ResetForm();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
				oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_SD050_CreateItems
		/// </summary>
		private void PS_SD050_CreateItems()
		{
			try
			{
				oDS_PS_SD050H = oForm.DataSources.DBDataSources.Item("@PS_SD050H");
				oDS_PS_SD050L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("DocDateF", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.DataSources.UserDataSources.Add("DocDateT", SAPbouiCOM.BoDataType.dt_DATE, 8);

				oForm.Items.Item("DocDateF").Specific.DataBind.SetBound(true, "", "DocDateF");
				oForm.Items.Item("DocDateT").Specific.DataBind.SetBound(true, "", "DocDateT");

				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD050_SetComboBox
		/// </summary>
		private void PS_SD050_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("Gbn").Specific.ValidValues.Add("0", "기성매출");
				oForm.Items.Item("Gbn").Specific.ValidValues.Add("1", "기성매출정산");
				oForm.Items.Item("Gbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("SGbn").Specific.ValidValues.Add("0", "기성매출");
				oForm.Items.Item("SGbn").Specific.ValidValues.Add("1", "기성매출정산");
				oForm.Items.Item("SGbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				oMat.Columns.Item("Gbn").ValidValues.Add("0", "기성매출");
				oMat.Columns.Item("Gbn").ValidValues.Add("1", "기성매출정산");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD050_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD050_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_SD050L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD050L.Offset = oRow;
				oDS_PS_SD050L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD050_LoadCaption
		/// </summary>
		private void PS_SD050_LoadCaption()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "추가";
					oForm.Items.Item("Btn_del").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "수정";
					oForm.Items.Item("Btn_del").Enabled = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD050_EnableFormItem
		/// </summary>
		private void PS_SD050_EnableFormItem()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("ItemCode").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("ItemCode").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD050_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD050_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
				switch (oUID)
				{
					case "ItemCode":
						sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

						sQry = "    Select          a.CardCode, ";
						sQry += "                   a.CardName,";
						sQry += "                   a.DocEntry, ";
						sQry += "                   b.LineNum ";
						sQry += " From           ORDR a, ";
						sQry += "                   RDR1 b ";
						sQry += " Where          a.DocEntry = b.DocEntry ";
						sQry += "                   and a.Canceled = 'N' ";
						sQry += "                   and b.ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("CardCode").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
						oForm.Items.Item("ORDRNum").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
						oForm.Items.Item("RDR1Num").Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_SD050_ResetForm
		/// </summary>
		private void PS_SD050_ResetForm()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Items.Item("ItemCode").Specific.Value = "";
				oForm.Items.Item("ItemName").Specific.Value = "";
				oForm.Items.Item("CardCode").Specific.Value = "";
				oForm.Items.Item("CardName").Specific.Value = "";
				oForm.Items.Item("Qty").Specific.Value = 0;
				oForm.Items.Item("Amt").Specific.Value = 0;
				oForm.Items.Item("Comments").Specific.Value = "";

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_SD050H]";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_SD050_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_SD050_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
                {
					errMessage = "제품코드(작번)은 필수사항입니다. 확인하세요.";
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return returnValue;
		}

		/// <summary>
		/// PS_SD050_LoadData
		/// </summary>
		private void PS_SD050_LoadData()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string SItemCode;
			string DocDateT;
			string DocDateF;
			string SGbn;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				SItemCode = oForm.Items.Item("SItemCode").Specific.Value.ToString().Trim();
				DocDateF  = oForm.Items.Item("DocDateF").Specific.Value.ToString().Trim();
				DocDateT  = oForm.Items.Item("DocDateT").Specific.Value.ToString().Trim();
				SGbn      = oForm.Items.Item("SGbn").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(SItemCode))
                {
					SItemCode = "%";
				}

				if (string.IsNullOrEmpty(DocDateF))
                {
					DocDateF = "19000101";
				}
					
				if (string.IsNullOrEmpty(DocDateT))
                {
					DocDateT = "20991231";
				}
					
				sQry = "EXEC [PS_SD050_01] '" + SGbn + "','" + SItemCode + "','" + DocDateF + "','" + DocDateT + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_SD050L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_SD050_AddMatrixRow(0, true);
					PS_SD050_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				oForm.Freeze(true);

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_SD050L.Size)
					{
						oDS_PS_SD050L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_SD050L.Offset = i;

					oDS_PS_SD050L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_SD050L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item(1).Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_SD050L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColNum02", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColNum03", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(10).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(11).Value.ToString().Trim());
					oDS_PS_SD050L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(12).Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);				
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD050_DeleteData
		/// </summary>
		private void PS_SD050_DeleteData()
		{
			string DocEntry;
			string sQry;
			string errMessage = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

					sQry = "Select Count(*) From [@PS_SD050H] where DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "Delete From [@PS_SD050H] where DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PS_SD050_ResetForm();
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.Items.Item("Btn_ret").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_SD050_UpdateData
		/// 데이타 insert
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_SD050_UpdateData(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool returnValue = false;
			string sQry;
			string errMessage = string.Empty;

			string CardCode;
			string BPLId;
			string DocEntry;
			string ItemCode;
			string ItemName;
			string Gbn;
			string RDR1Num;
			string CardName;
			string ORDRNum;
			string Comments;
			string DocDate;
			decimal Qty;
			decimal Amt;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLId    = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate  = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
				ORDRNum  = oForm.Items.Item("ORDRNum").Specific.Value.ToString().Trim();
				RDR1Num  = oForm.Items.Item("RDR1Num").Specific.Value.ToString().Trim();
				Gbn      = oForm.Items.Item("Gbn").Specific.Value.ToString().Trim();
				Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
				Qty = Convert.ToDecimal(oForm.Items.Item("Qty").Specific.Value.ToString().Trim());
				Amt = Convert.ToDecimal(oForm.Items.Item("Amt").Specific.Value.ToString().Trim());

				if (string.IsNullOrEmpty(DocEntry))
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!";
					throw new Exception();
				}

				sQry = "Update [@PS_SD050H]";
				sQry += " set ";
				sQry += " U_BPLId = '" + BPLId + "',";
				sQry += " U_ItemCode = '" + ItemCode + "',";
				sQry += " U_ItemName = '" + ItemName + "',";
				sQry += " U_CardCode = '" + CardCode + "',";
				sQry += " U_CardName  = '" + CardName + "',";
				sQry += " U_ORDRNum  = '" + ORDRNum + "',";
				sQry += " U_RDR1Num  = '" + RDR1Num + "',";
				sQry += " U_Gbn  = '" + Gbn + "',";
				sQry += " U_Qty = '" + Qty + "',";
				sQry += " U_Amt = '" + Amt + "',";
				sQry += " U_Comments = '" + Comments + "'";
				sQry += " Where DocEntry = '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				returnValue = true;
				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_SD050_AddPurchaseDemand
		/// 데이타 insert
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_SD050_AddPurchaseDemand(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool returnValue = false;
			string sQry;

			string CardCode;
			string ItemCode;
			string DocEntry;
			string BPLId;
			string ItemName;
			string Gbn;
			string ORDRNum;
			string CardName;
			string RDR1Num;
			object Comments;
			string DocDate;
			decimal Qty;
			decimal Amt;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLId    = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate  = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
				ORDRNum  = oForm.Items.Item("ORDRNum").Specific.Value.ToString().Trim();
				RDR1Num  = oForm.Items.Item("RDR1Num").Specific.Value.ToString().Trim();
				Gbn      = oForm.Items.Item("Gbn").Specific.Value.ToString().Trim();
				Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
				Qty = Convert.ToDecimal(oForm.Items.Item("Qty").Specific.Value.ToString().Trim());
				Amt = Convert.ToDecimal(oForm.Items.Item("Amt").Specific.Value.ToString().Trim());

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_SD050H]";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = Convert.ToString(1);
				}
				else
				{
					DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
				}

				sQry = "INSERT INTO [@PS_SD050H]";
				sQry +=  " (";
				sQry +=  " DocEntry,";
				sQry +=  " DocNum,";
				sQry +=  " U_BPLId,";
				sQry +=  " U_DocDate,";
				sQry +=  " U_ItemCode,";
				sQry +=  " U_ItemName,";
				sQry +=  " U_CardCode,";
				sQry +=  " U_CardName,";
				sQry +=  " U_Qty,";
				sQry +=  " U_Amt,";
				sQry +=  " U_ORDRNum,";
				sQry +=  " U_RDR1Num,";
				sQry +=  " U_Comments,";
				sQry +=  " U_Gbn";
				sQry +=  " ) ";
				sQry +=  "VALUES(";
				sQry +=  DocEntry + ",";
				sQry +=  DocEntry + ",";
				sQry +=  "'" + BPLId + "',";
				sQry +=  "'" + DocDate + "',";
				sQry +=  "'" + ItemCode + "',";
				sQry +=  "'" + ItemName + "',";
				sQry +=  "'" + CardCode + "',";
				sQry +=  "'" + CardName + "',";
				sQry +=  "'" + Qty + "',";
				sQry +=  "'" + Amt + "',";
				sQry +=  "'" + ORDRNum + "',";
				sQry +=  "'" + RDR1Num + "',";
				sQry +=  "'" + Comments + "',";
				sQry +=  "'" + Gbn + "'";
				sQry +=  ")";
				oRecordSet.DoQuery(sQry);

				returnValue = true;
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn_save")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD050_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_SD050_AddPurchaseDemand(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_SD050_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD050_LoadCaption();
							PS_SD050_LoadData();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_SD050_UpdateData(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_SD050_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD050_LoadCaption();
							PS_SD050_LoadData();
							oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}
					else if (pVal.ItemUID == "Btn_ret")
					{
						PS_SD050_LoadData();
					}
					else if (pVal.ItemUID == "Btn_del")
					{
						PS_SD050_DeleteData();
						PS_SD050_LoadData();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
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
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
							oMat.SelectRow(pVal.Row, true, false);
							oForm.Items.Item("DocEntry").Specific.Value = oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("DocDate").Specific.Value = oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("ItemCode").Specific.Value = oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Qty").Specific.Value = oMat.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Amt").Specific.Value = oMat.Columns.Item("Amt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Comments").Specific.Value = oMat.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Gbn").Specific.Select(oMat.Columns.Item("Gbn").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_SD050_LoadCaption();
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
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			string DocDate;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "ItemCode")
						{
							PS_SD050_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						if (pVal.ItemUID == "DocDate")
						{
							DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

							sQry = "Select left('" + DocDate + "',6) + '01', Convert(char(8),Dateadd(dd, -1, left(convert(char(8),Dateadd(mm, 1, '" + DocDate + "'),112), 6) + '01'),112)";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("DocDateF").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("DocDateT").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD050H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD050L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_SD050H.RemoveRecord(oDS_PS_SD050H.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_SD050_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_SD050H.GetValue("U_CntcCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_SD050_AddMatrixRow(oMat.RowCount, false);
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
		/// Raise_RightClickEvent
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
							PS_SD050_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_SD050_LoadCaption();
							oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
							break;
						case "7169": //엑셀 내보내기
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
