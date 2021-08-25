using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 포장생산 작업지시서 발행
	/// </summary>
	internal class PS_PP015 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP015H;  //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP015L;  //등록라인
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private SAPbouiCOM.BoFormMode oLast_Mode;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP015.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP015_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP015");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);

				PS_PP015_CreateItems();
				PS_PP015_SetComboBox();
				PS_PP015_Initialize();
				PS_PP015_ClearForm();
				PS_PP015_AddMatrixRow(0, true);

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
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
		/// PS_PP015_CreateItems
		/// </summary>
		private void PS_PP015_CreateItems()
		{
			try
			{
				oDS_PS_PP015H = oForm.DataSources.DBDataSources.Item("@PS_PP015H");
				oDS_PS_PP015L = oForm.DataSources.DBDataSources.Item("@PS_PP015L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

				oDS_PS_PP015H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_SetComboBox
		/// </summary>
		private void PS_PP015_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				oForm.Items.Item("StdPoYN").Specific.ValidValues.Add("Y", "PO정식발행");
				oForm.Items.Item("StdPoYN").Specific.ValidValues.Add("N", "기PO발행분저장");
				oForm.Items.Item("StdPoYN").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품목대분류
				sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("ItmBsort").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//품목중분류
				sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Order by U_Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("ItmMsort").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
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
		/// PS_PP015_Initialize
		/// </summary>
		private void PS_PP015_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("Send").Specific.Value = "영업";
				oForm.Items.Item("Receive").Specific.Value = "생산";
				PS_PP015_NumberSet();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_ClearForm
		/// </summary>
		private void PS_PP015_ClearForm()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP015'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oForm.Items.Item("DocNum").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocNum").Specific.Value = DocNum;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP015_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP015L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP015L.Offset = oRow;
				oDS_PS_PP015L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_EnableFormItem
		/// </summary>
		private void PS_PP015_EnableFormItem()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP015_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				switch (oUID)
				{
					case "Mat01":
						if (oCol == "ReqNum")
						{
							if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("ReqNum").Cells.Item(oRow).Specific.Value.ToString().Trim()))
							{
								oMat.FlushToDataSource();
								PS_PP015_AddMatrixRow(oMat.RowCount, false);
								oMat.Columns.Item("ReqNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
							sQry = " select OrdrNum = Convert(Nvarchar(10),t.DocNum) + '-' + Convert(Nvarchar(10),t1.LineNum), ";
							sQry += " t1.ItemCode, ItemName = t1.Dscription, t2.U_ItmBsort, t2.U_ItmMsort, ";
							sQry += " t.DocNum, t1.LineNum, t1.Quantity, t.DocDate, DueDate = t1.ShipDate, ";
							sQry += " t.CardCode , t.CardName ";
							sQry += " from ORDR t Inner Join RDR1 t1 On t.DocEntry = t1.DocEntry and t1.LineStatus = 'O' ";
							sQry += " Inner Join OITM t2 On t1.ItemCode = t2.ItemCode ";
							sQry += " where Convert(Nvarchar(10),t.DocNum) + '-' + Convert(Nvarchar(10),t1.LineNum) = '" + oMat.Columns.Item("ReqNum").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount == 0)
							{
								errMessage = "조회 결과가 없습니다. 확인하세요.";
								throw new Exception();
							}

							oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim();
							oMat.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
							oMat.Columns.Item("ItmBsort").Cells.Item(oRow).Specific.Select(oRecordSet.Fields.Item("U_ItmBsort").Value.ToString().Trim());
							oMat.Columns.Item("ItmMsort").Cells.Item(oRow).Specific.Select(oRecordSet.Fields.Item("U_ItmMsort").Value.ToString().Trim());
							oMat.Columns.Item("SjDocNum").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("DocNum").Value.ToString().Trim();
							oMat.Columns.Item("SjLinNum").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("LineNum").Value.ToString().Trim();
							oMat.Columns.Item("SjQty").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("Quantity").Value.ToString().Trim();
							oMat.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd");
							oMat.Columns.Item("SjDcDate").Cells.Item(oRow).Specific.Value = Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd");
							oMat.Columns.Item("SjDuDate").Cells.Item(oRow).Specific.Value = Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd");
							oMat.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("CardCode").Value.ToString().Trim();
							oMat.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("CardName").Value.ToString().Trim();

							oMat.Columns.Item("ReqNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (oCol == "ItemCode")
						{
							if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
							{
								oMat.FlushToDataSource();
								PS_PP015_AddMatrixRow(oMat.RowCount, false);
								oMat.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
							sQry = " select t.ItemCode, t.ItemName, t.U_ItmBsort, t.U_ItmMsort ";
							sQry += " from OITM t ";
							sQry += " where t.ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount == 0)
							{
								errMessage = "조회 결과가 없습니다. 확인하세요.";
								throw new Exception();
							}

							oMat.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
							oMat.Columns.Item("ItmBsort").Cells.Item(oRow).Specific.Select(oRecordSet.Fields.Item("U_ItmBsort").Value.ToString().Trim());
							oMat.Columns.Item("ItmMsort").Cells.Item(oRow).Specific.Select(oRecordSet.Fields.Item("U_ItmMsort").Value.ToString().Trim());
							oMat.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (oCol == "CardCode")
						{
							sQry = "Select CardName from OCRD Where CardCode = '" + oMat.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount == 0)
							{
								oMat.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = "";
								errMessage = "거래처명 조회 결과가 없습니다. 확인하세요.";
								throw new Exception();
							}

							oMat.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("CardName").Value.ToString().Trim();
							oMat.Columns.Item("CardCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						break;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP015_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP015_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_PP015H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP015H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					errMessage = "지시일은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				functionReturnValue = true;
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
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP015_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP015_DelMatrixSpaceLine()
		{
			bool functionReturnValue = false;

			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				//라인
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}
				else if (oMat.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_PP015L.GetValue("U_ReqNum", 0).ToString().Trim()))
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요..";
					throw new Exception();
				}

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP015L.GetValue("U_ItemCode", i).ToString().Trim()))
					{
						errMessage = "품목코드는 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (Convert.ToDouble(oDS_PS_PP015L.GetValue("U_DueDate", i).ToString().Trim()) == 0)
					{
						errMessage = "납기일은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP015L.GetValue("U_CardCode", i).ToString().Trim()))
					{
						errMessage = "거래처는 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
				}
				oMat.LoadFromDataSource();
				functionReturnValue = true;
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
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP015_DeleteEmptyRow
		/// </summary>
		private void PS_PP015_DeleteEmptyRow()
		{
			int i;

			try
			{
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP015L.GetValue("U_ItemCode", i).ToString().Trim()))
					{
						oDS_PS_PP015L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
					}
				}
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP015_NumberSet
		/// </summary>
		private void PS_PP015_NumberSet()
		{
			string BPLID;
			string YM;
			string Cnt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				YM = codeHelpClass.Left(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim(), 6);

				sQry = " Select right('0' + Convert(Nvarchar(2), Convert(integer,Right(Max(Isnull(U_Number,'00')),2)) + 1),2) ";
				sQry += " From [@PS_PP015H] ";
				sQry += " Where U_BPLId = '" + BPLID + "' And Convert(char(6),U_DocDate,112) = '" + YM + "' and Isnull(U_Number,'') <> '' and Canceled = 'N' ";
				oRecordSet.DoQuery(sQry);

				Cnt = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				if (string.IsNullOrEmpty(Cnt))
				{
					Cnt = "01";
				}
				oForm.Items.Item("Number").Specific.Value = YM + "-" + Cnt;
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
		/// PS_PP015_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP015_PrintReport()
		{
			string WinTitle;
			string ReportName;
			string DocNum;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocNum = oDS_PS_PP015H.GetValue("DocNum", 0).ToString().Trim();

				WinTitle = "[PS_PP015] 생산지시서";
				ReportName = "PS_PP015_01.RPT";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocNum));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP015_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP015_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.Items.Item("StdPoYN").Specific.Value.ToString().Trim() == "Y")
							{
								PS_PP015_NumberSet();
							}
							else
							{
								oForm.Items.Item("Number").Specific.Value = "";
							}

							oMat.FlushToDataSource();
							oMat.LoadFromDataSource();
							PS_PP015_DeleteEmptyRow();
							oLast_Mode = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							oLast_Mode = oForm.Mode;

							if (PS_PP015_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP015_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
						{
							oLast_Mode = oForm.Mode;
						}
					}
					else if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP015_PrintReport);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
							{
								PS_PP015_AddMatrixRow(oMat.RowCount, false);
								oLast_Mode = (BoFormMode)100;
							}
							else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
							{
								PS_PP015_AddMatrixRow(oMat.RowCount, false);
								PS_PP015_EnableFormItem();
								oLast_Mode = (BoFormMode)100;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
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
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ReqNum")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("ReqNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							else if (pVal.ColUID == "ItemCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							else if (pVal.ColUID == "CardCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "StdPoYN")
					{
						if (oForm.Items.Item("StdPoYN").Specific.Value.ToString().Trim() == "N")
						{
							oForm.Items.Item("Number").Specific.Value = "";
						}
						else
						{
							PS_PP015_NumberSet();
						}
					}
					if (pVal.ItemUID == "BPLId")
					{
						PS_PP015_NumberSet();
					}
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oLast_Item_UID = pVal.ItemUID;
							oLast_Col_UID = pVal.ColUID;
							oLast_Col_Row = pVal.Row;
							oMat.SelectRow(pVal.Row, true, false);
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
						if (pVal.ItemUID == "DocDate")
						{
							PS_PP015_NumberSet();
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ReqNum" || pVal.ColUID == "ItemCode" || pVal.ColUID == "CardCode")
							{
								PS_PP015_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
				oForm.Freeze(false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP015H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP015L);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
						case "1285": //복원
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
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
							PS_PP015_EnableFormItem();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1286": //닫기
							break;
						case "1281": //찾기
							PS_PP015_EnableFormItem();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							break;
						case "1282": //추가
							oDS_PS_PP015H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							PS_PP015_Initialize();
							PS_PP015_EnableFormItem();
							PS_PP015_ClearForm();
							PS_PP015_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_PP015_EnableFormItem();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("ReqNum").Cells.Item(oMat.VisualRowCount).Specific.Value.ToString().Trim()))
								{
									PS_PP015_AddMatrixRow(oMat.RowCount, false);
								}
							}
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
								}
								oMat.FlushToDataSource();
								oDS_PS_PP015L.RemoveRecord(oDS_PS_PP015L.Size - 1); // Mat01에 마지막라인(빈라인) 삭제
								oMat.Clear();
								oMat.LoadFromDataSource();
							}
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
