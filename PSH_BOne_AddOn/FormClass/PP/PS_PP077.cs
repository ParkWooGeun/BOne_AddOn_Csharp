using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 휘팅서울포장등록
	/// </summary>
	internal class PS_PP077 : PSH_BaseClass
	{
		private string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat;
		public SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DBDataSource oDS_PS_PP077H; //등록헤더
		private string VIDocNum;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP077.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP077_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP077");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP077_CreateItems();
				PS_PP077_SetComboBox();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_PP077_CreateItems
		/// </summary>
		private void PS_PP077_CreateItems()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oDS_PS_PP077H = oForm.DataSources.DBDataSources.Item("@PS_PP077H");
				oMat = oForm.Items.Item("Mat01").Specific;
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("ZTEMP");

				oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");
				oForm.DataSources.UserDataSources.Item("FrDate").Value = DateTime.Now.ToString("yyyyMM01");

				oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");
				oForm.DataSources.UserDataSources.Item("ToDate").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
				oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

				//작업자 기본설정
				oForm.Items.Item("EmpId").Specific.Value = dataHelpClass.User_MSTCOD();

				oMat.Columns.Item("NPkQty").Editable = true;
				oMat.Columns.Item("NPkWt").Editable = true;
				oMat.Columns.Item("InDate").Editable = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP077_SetComboBox
		/// </summary>
		private void PS_PP077_SetComboBox()
		{
			try
			{
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("1", "포장대기");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("2", "포장완료");
				oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("MovDocNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP077_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP077_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;
			double SumNPkQty = 0 ; //포장수량합계
			double SumNPkWt = 0;   //포장중량합계
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				switch (oUID)
				{
					case "ItemCode":
						sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItemName").Specific.String = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;

					case "EmpId":
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("EmpId").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("EmpName").Specific.String = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;

					case "DocDate":
						oMat.FlushToDataSource();
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oDS_PS_PP077H.Offset = i;
							oDS_PS_PP077H.SetValue("U_InDate", i, Convert.ToDateTime(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()).ToString("yyyyMMdd"));
						}
						oMat.LoadFromDataSource();
						break;
				}
				
				if (oUID == "Mat01")  //Line
				{
					switch (oCol)
					{
						case "NPkQty":
						case "NPkWt":
							oMat.FlushToDataSource();

							for (i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								oDS_PS_PP077H.Offset = i;
								SumNPkQty += Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim());
								SumNPkWt += Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim());
							}
							oForm.Items.Item("SumNPkQty").Specific.Value = SumNPkQty;
							oForm.Items.Item("SumNPkWt").Specific.Value = SumNPkWt;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP077_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP077_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
                {
					errMessage = "입고전기일은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("EmpId").Specific.Value.ToString().Trim()))
				{
					errMessage = "작업자코드는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if ( string.IsNullOrEmpty(oForm.Items.Item("EmpName").Specific.Value.ToString().Trim()))
				{
					errMessage = "작업자명이 없습니다. 작업자코드를 확인하여 주십시오.";
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
		/// PS_PP077_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP077_DelMatrixSpaceLine()
		{
			bool functionReturnValue = false;

			int i;
			int j = 0;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하여 주십시오.";
					throw new Exception();
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						if (oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() == "Y")
						{
							oDS_PS_PP077H.Offset = i;
							if (string.IsNullOrEmpty(oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim()) || Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim()) == 0)
							{
								errMessage = "포장수량이 0 또는 없습니다. 확인하여 주십시오.";
								throw new Exception();
							}
							else if (string.IsNullOrEmpty(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim()) | Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim()) == 0)
							{
								errMessage = "포장중량이 0 또는 없습니다. 확인하여 주십시오.";
								throw new Exception();
							}

							j += 1;	//체크된 라인Count
						}
					}
				}
				// 체크된 라인Count 가
				if (j == 0)
				{
					errMessage = "선택되어진 라인이 없습니다. 확인하여 주십시오.";
					throw new Exception();
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
		/// PS_PP077_ResizeForm
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void PS_PP077_ResizeForm(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Items.Item("Grid01").Top = 40;
				oForm.Items.Item("Grid01").Height = (oForm.Height / 2) - 75;
				oForm.Items.Item("Grid01").Left = 6;
				oForm.Items.Item("Grid01").Width = oForm.Width - 21;

				oForm.Items.Item("DistrQty").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("Distribute").Top = (oForm.Height / 2) - 37;

				oForm.Items.Item("s9").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("s9").Left = (oForm.Width / 2) - 65;

				oForm.Items.Item("EmpId").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("EmpId").Left = (oForm.Width / 2) + 16;

				oForm.Items.Item("EmpName").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("EmpName").Left = (oForm.Width / 2) + 116;

				oForm.Items.Item("s8").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("s8").Left = (oForm.Width) - 190;

				oForm.Items.Item("DocDate").Top = (oForm.Height / 2) - 32;
				oForm.Items.Item("DocDate").Left = (oForm.Width) - 112;

				oForm.Items.Item("Mat01").Top = (oForm.Height / 2) - 15;
				oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 38;
				oForm.Items.Item("Mat01").Left = 6;
				oForm.Items.Item("Mat01").Width = oForm.Width - 21;

				if (oGrid.Rows.Count > 0)
				{
					oGrid.AutoResizeColumns();
				}
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP077_SearchGridData
		/// </summary>
		private void PS_PP077_SearchGridData()
		{
			string sQry;
			string MovDocNo;
			string FrDate;
			string ToDate;
			string ItemCode;
			string Status;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				oForm.Items.Item("DistrQty").Specific.Value = "";
				oForm.Items.Item("SumPkQty").Specific.Value = "";
				oForm.Items.Item("SumPkWt").Specific.Value = "";
				oForm.Items.Item("SumNPkQty").Specific.Value = "";
				oForm.Items.Item("SumNPkWt").Specific.Value = "";
				oForm.Items.Item("EmpId").Specific.Value = dataHelpClass.User_MSTCOD();
				PS_PP077_FlushToItemValue("EmpId", 0, "");
				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oMat.Clear();

				MovDocNo = oForm.Items.Item("MovDocNo").Specific.Value.ToString().Trim();
				if (string.IsNullOrEmpty(MovDocNo))
                {
					MovDocNo = "%";
				}

				FrDate = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
				ToDate = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
				if (string.IsNullOrEmpty(FrDate))
                {
					FrDate = "1900-01-01";
				}
				if (string.IsNullOrEmpty(ToDate))
                {
					ToDate = "2100-12-31";
				}
					
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				if (string.IsNullOrEmpty(ItemCode))
                {
					ItemCode = "%";
				}

				Status = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

				sQry = "EXEC PS_PP077_01 '" + MovDocNo + "','" + FrDate + "','" + ToDate + "','" + ItemCode + "','" + Status + "'";

				// Procedure 실행(Grid 사용)
				oForm.DataSources.DataTables.Item(0).ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("ZTEMP");

				PS_PP077_SetGrid();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP077_SetGrid Grid 꾸며주기
		/// </summary>
		private void PS_PP077_SetGrid()
		{
			short i;
			string sColsTitle;

			try
			{
				oForm.Freeze(true);
				oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(2)).LinkedObjectType = "4";

				for (i = 0; i <= oGrid.Columns.Count - 1; i++)
				{
					sColsTitle = oGrid.Columns.Item(i).TitleObject.Caption;

					oGrid.Columns.Item(i).Editable = false;

					if (oGrid.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
					{
						oGrid.Columns.Item(i).RightJustified = true;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP077_SearchMatrixData
		/// </summary>
		private void PS_PP077_SearchMatrixData()
		{
			string sQry;

			int i;
			int j;
			int Cnt;
			double SumPkQty = 0;  //이동요청수량합계
			double SumPkWt = 0;   //이동요청중량합계
			double SumNPkWt = 0;  //포장중량합계
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				for (i = 0; i <= oGrid.Rows.Count - 1; i++)
				{
					if (oGrid.Rows.IsSelected(i) == true)
					{
						sQry = "EXEC PS_PP077_02 '" + oGrid.DataTable.GetValue(0, i).ToString().Trim() + "', '" + oGrid.DataTable.GetValue(1, i).ToString().Trim() + "', '" + oGrid.DataTable.GetValue(2, i).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						Cnt = oDS_PS_PP077H.Size;
						if (Cnt > 0)
						{
							for (j = 0; j <= Cnt - 1; j++)
							{
								oDS_PS_PP077H.RemoveRecord(oDS_PS_PP077H.Size - 1);
							}
							if (Cnt == 1)
							{
								oDS_PS_PP077H.Clear();
							}
						}
						oMat.LoadFromDataSource();

						j = 1;
						while (!oRecordSet.EoF)
						{
							if (oDS_PS_PP077H.Size < j)
							{
								oDS_PS_PP077H.InsertRecord(j - 1); //라인추가
							}
							oDS_PS_PP077H.SetValue("DocNum", j - 1, Convert.ToString(j));
							oDS_PS_PP077H.SetValue("U_Check", j - 1, "Y");
							oDS_PS_PP077H.SetValue("U_MovDocNo", j - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PP070No", j - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PP070NoL", j - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PorNum", j - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_ItemCode", j - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_ItemName", j - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PkQty", j - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PkWt", j - 1, oRecordSet.Fields.Item(7).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_NPkWt", j - 1, oRecordSet.Fields.Item(8).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_OPkQty", j - 1, oRecordSet.Fields.Item(9).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_OPkWt", j - 1, oRecordSet.Fields.Item(10).Value.ToString().Trim());
							//포장대기
							if (oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim() == "1")
							{
								oDS_PS_PP077H.SetValue("U_InDate", j - 1, DateTime.Now.ToString("yyyyMMdd")); //포장완료
							}
							else
							{ 
				                oDS_PS_PP077H.SetValue("U_InDate", j - 1, Convert.ToDateTime(oRecordSet.Fields.Item(11).Value.ToString().Trim()).ToString("yyyyMMdd"));
							}

							SumPkQty = Convert.ToDouble(oGrid.DataTable.GetValue(7, i).ToString().Trim());
							SumPkWt = Convert.ToDouble(oGrid.DataTable.GetValue(8, i).ToString().Trim());
							SumNPkWt += Convert.ToDouble(oRecordSet.Fields.Item(8).Value.ToString().Trim()); //포장중량합계
							j += 1;
							oRecordSet.MoveNext();
						}
						oMat.LoadFromDataSource();
					}
				}
				oForm.Items.Item("SumPkQty").Specific.Value = SumPkQty;
				oForm.Items.Item("SumPkWt").Specific.Value = SumPkWt;
				oForm.Items.Item("SumNPkQty").Specific.Value = 0;
				oForm.Items.Item("SumNPkWt").Specific.Value = SumNPkWt;
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
		/// PS_PP077_DistiributeData
		/// 포장중량을 기준으로 포장수량 전량배부
		/// </summary>
		private void PS_PP077_DistiributeData()
		{
			int i;
			double SumNPkWt;
			double SumNPkQty = 0;
			double DistrQty;

			try
			{
				oForm.Freeze(true);

				oMat.FlushToDataSource();
				if (oMat.VisualRowCount > 0)
				{
					SumNPkWt = Convert.ToDouble(oForm.Items.Item("SumNPkWt").Specific.Value.ToString().Trim());
					DistrQty = Convert.ToDouble(oForm.Items.Item("DistrQty").Specific.Value.ToString().Trim());

					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						oDS_PS_PP077H.Offset = i;
						oDS_PS_PP077H.SetValue("U_NPkQty", i, Convert.ToString(System.Math.Round(DistrQty * (Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim()) / SumNPkWt), 0)));

						SumNPkQty += System.Math.Round(DistrQty * (Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim()) / SumNPkWt), 0);
					}
					oForm.Items.Item("SumNPkQty").Specific.Value = SumNPkQty;
				}
				oMat.LoadFromDataSource();
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
		/// PS_PP077_InsertoInventoryGenEntry
		/// </summary>
		/// <param name="ChkType"></param>
		/// <returns></returns>
		private bool PS_PP077_InsertoInventoryGenEntry(int ChkType)
		{
			bool functionReturnValue = false;

			//입고 DI
			SAPbobsCOM.Documents DI_oInventoryGenEntry = null; //재고입고 문서 객체
			int RetVal;
			int i;
			int oRow;
			int ErrCode;
			string ErrMsg = string.Empty;
			string oDate;

			try
			{
				//입고
				PSH_Globals.oCompany.StartTransaction();
				DI_oInventoryGenEntry = null;
				DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry); // 문서타입(입고)

				i = 1;
				var _with1 = DI_oInventoryGenEntry;

				// Header
				oDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				oDate = Convert.ToDateTime(oDate.Substring(0, 4) + "-" + oDate.Substring(4, 2) + "-" + oDate.Substring(6, 2)).ToString("yyyy-MM-dd");
				//_with1.DocDate = Convert.ToDateTime(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()).ToString("yyyy-MM-dd");
				_with1.DocDate = Convert.ToDateTime(oDate);
				//_with1.TaxDate = Convert.ToDateTime(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()).ToString("yyyy-MM-dd");
				_with1.TaxDate = Convert.ToDateTime(oDate);
				_with1.Comments = "포장처리등록 입고";

				// Line
				for (oRow = 0; oRow <= oDS_PS_PP077H.Size - 1; oRow++)
				{
					if (oDS_PS_PP077H.GetValue("U_Check", oRow).ToString().Trim() == "Y")
					{
						if (_with1.Lines.Count < i)
						{
							_with1.Lines.Add();
						}

						_with1.Lines.SetCurrentLine(i - 1);
						_with1.Lines.ItemCode = oDS_PS_PP077H.GetValue("U_ItemCode", oRow).ToString().Trim();
						_with1.Lines.ItemDescription = oDS_PS_PP077H.GetValue("U_ItemName", oRow).ToString().Trim();
						_with1.Lines.Quantity = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", oRow).ToString().Trim());
						_with1.Lines.WarehouseCode = "104";	//제품-서울
						i += 1;
					}
				}

				// 입고전송
				RetVal = DI_oInventoryGenEntry.Add();
				if (0 != RetVal)
				{
					PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
					throw new Exception();
				}
				if (ChkType != 2)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				else
				{
					PSH_Globals.oCompany.GetNewObjectCode(out VIDocNum);

					//[@PS_PP077H]에 Insert
					if (PS_PP077_SaveData("B", false) == false)
					{
						throw new Exception();
					}
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrMsg != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(ErrMsg);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP077_InsertoInventoryGenExit
		/// </summary>
		/// <param name="ChkType"></param>
		/// <returns></returns>
		private bool PS_PP077_InsertoInventoryGenExit(int ChkType)
		{
			bool functionReturnValue = false;
			//출고 DI
			SAPbobsCOM.Documents DI_oInventoryGenExit = null; //재고입고 문서 객체
			int RetVal;
			int ErrCode;
			string ErrMsg = string.Empty;
			int i;
			int oRow;
			string oDate;

			try
			{
				//출고
				PSH_Globals.oCompany.StartTransaction();
				DI_oInventoryGenExit = null;
				DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); // 문서타입(입고)

				i = 1;
				var _with2 = DI_oInventoryGenExit;

				// Header
				//_with2.DocDate = Convert.ToDateTime(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()).ToString("yyyy-MM-dd");
				//_with2.TaxDate = Convert.ToDateTime(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()).ToString("yyyy-MM-dd");

				oDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				oDate = Convert.ToDateTime(oDate.Substring(0, 4) + "-" + oDate.Substring(4, 2) + "-" + oDate.Substring(6, 2)).ToString("yyyy-MM-dd");
				_with2.DocDate = Convert.ToDateTime(oDate);
				_with2.TaxDate = Convert.ToDateTime(oDate);
				_with2.Comments = "포장처리등록 입고";

				// Line
				for (oRow = 0; oRow <= oDS_PS_PP077H.Size - 1; oRow++)
				{
					if (oDS_PS_PP077H.GetValue("U_Check", oRow).ToString().Trim() == "Y")
					{
						if (_with2.Lines.Count < i)
						{
							_with2.Lines.Add();
						}
						_with2.Lines.SetCurrentLine(i - 1);
						_with2.Lines.ItemCode = oDS_PS_PP077H.GetValue("U_ItemCode", oRow).ToString().Trim();
						_with2.Lines.ItemDescription = oDS_PS_PP077H.GetValue("U_ItemName", oRow).ToString().Trim();
						_with2.Lines.Quantity = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", oRow).ToString().Trim());
						_with2.Lines.WarehouseCode = "104"; //제품-서울
						i += 1;
					}
				}

				// 출고전송
				RetVal = DI_oInventoryGenExit.Add();
				if (0 != RetVal)
				{
					PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
					throw new Exception();
				}
				if (ChkType != 2)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				else
				{
					// [@PS_PP077H]에 -수량으로 Insert
					if (PS_PP077_SaveData("B", true) == false)
					{
						throw new Exception();
					}
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrMsg != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(ErrMsg);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP077_SaveData
		/// </summary>
		/// <param name="A_B"></param>
		/// <param name="Cancel"></param>
		/// <returns></returns>
		private bool PS_PP077_SaveData(string A_B, bool Cancel)
		{
			bool functionReturnValue = false;

			string sQry;
			string errMessage = string.Empty;

			int i;
			int j;
			int DocNum;
			double PkWt;
			double NPkWt;
			double OPkWt = 0;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						oDS_PS_PP077H.Offset = i;
						if (oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() == "Y")
						{
							switch (A_B)
							{
								case "A":
									// 이동요청중량 보다 기포장중량+포장중량이 클수없음, Check!!!
									PkWt = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_PkWt", i).ToString().Trim());	//이동요청중량
									NPkWt = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim());	//포장중량
									
									for (j = 0; j <= oMat.VisualRowCount - 1; j++)
									{
										if (oDS_PS_PP077H.GetValue("U_Check", j).ToString().Trim() == "Y")
										{
											OPkWt += Convert.ToDouble(oDS_PS_PP077H.GetValue("U_OPkWt", j).ToString().Trim()); //기포장중량
										}
									}

									if (PkWt < NPkWt + OPkWt)
									{
										oMat.Columns.Item("NPkWt").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
										errMessage = "기포장중량 + 포장중량은 이동요청중량을 초과할 수 없습니다. 확인하여 주십시오.";
										throw new Exception();
									}
									break;

								case "B":
									// [@PS_PP077H]에 Insert
									sQry = "select max(DocNum) MaxDocNum from [@PS_PP077H]";
									oRecordSet.DoQuery(sQry);

									if (oRecordSet.RecordCount == 0)
									{
										DocNum = 1;
									}
									else
									{
										DocNum = Convert.ToInt32(oRecordSet.Fields.Item("MaxDocNum").Value.ToString().Trim()) + 1;
									}

									sQry = "insert into [@PS_PP077H] values (";
									sQry += "'" + DocNum + "','" + DocNum + "'";
									sQry += ",Null,0,Null,'N','N',Null,Null,Null,'N','O',Null,Null,Null,Null,Null,";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_MovDocNo", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PP070No", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PP070NoL", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PorNum", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_ItemCode", i).ToString().Trim() + "',";
									sQry += "'" + dataHelpClass.Make_ItemName(oDS_PS_PP077H.GetValue("U_ItemName", i).ToString().Trim()) + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PkQty", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PkWt", i).ToString().Trim() + "',";

									//취소 마이너스(-) insert
									if (Cancel == true)
									{
										sQry += "'" + -1 * Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim()) + "',";
										sQry += "'" + -1 * Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim()) + "',";
									}
									else
									{
										sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim() + "',";
										sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim() + "',";
									}
									sQry += "'" + oDS_PS_PP077H.GetValue("U_OPkQty", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_OPkWt", i).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_InDate", i).ToString().Trim() + "'";
									sQry += ",Null,Null,";
									sQry += "'" + VIDocNum + "',";
									sQry += "Null)";
									oRecordSet.DoQuery(sQry);

									// [@PS_PP040H](작업지시)에 Insert
									if (PS_PP077_AddPS_PP040(i, DocNum) == false)
									{
										throw new Exception();
									}
									break;
							}
						}
					}
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP077_AddPS_PP040
		/// </summary>
		/// <param name="i"></param>
		/// <param name="PP077HDocNum"></param>
		/// <returns></returns>
		private bool PS_PP077_AddPS_PP040(int i, int PP077HDocNum)
		{
			bool functionReturnValue = false;

			int j;
			string sQry;
			string errMessage = string.Empty;
			int AutoKey;
			string PP040H_DocEntry;

			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oMat.FlushToDataSource();

				if (oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() == "Y")
				{
					sQry = "EXEC [PS_PP077_03] '" + oDS_PS_PP077H.GetValue("U_MovDocNo", i).ToString().Trim() + "','" + oDS_PS_PP077H.GetValue("U_PP070No", i).ToString().Trim() + "','" + oDS_PS_PP077H.GetValue("U_PP070NoL", i).ToString().Trim() + "'";
					oRecordSet01.DoQuery(sQry);

					// DocEntry
					sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'";
					oRecordSet02.DoQuery(sQry);
					PP040H_DocEntry = oRecordSet02.Fields.Item("AutoKey").Value.ToString().Trim();
					AutoKey = Convert.ToInt32(PP040H_DocEntry) + 1;

					// Insert PS_PP040H
					sQry = "INSERT INTO [@PS_PP040H]";
					sQry += " (";
					sQry += " DocEntry,";
					sQry += " DocNum,";
					sQry += " Period,";
					sQry += " Series,";
					sQry += " Object,";
					sQry += " UserSign,";
					sQry += " CreateDate,";
					sQry += " CreateTime,";
					sQry += " U_OrdType,";
					sQry += " U_OrdGbn,";
					sQry += " U_BPLId,";
					sQry += " U_ItemCode,";
					sQry += " U_ItemName,";
					sQry += " U_OrdNum,";
					sQry += " U_OrdSub1,";
					sQry += " U_OrdSub2,";
					sQry += " U_PP030HNo,";
					sQry += " U_DocType,";
					sQry += " Canceled,";
					sQry += " U_DocDate";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += "'" + PP040H_DocEntry + "',";
					sQry += "'" + PP040H_DocEntry + "',";
					sQry += "'21',";
					sQry += "'-1',";
					sQry += "'PS_PP040',";
					sQry += "'1',";
					sQry += "'" + DateTime.Now.ToString("yyyyMMdd") + "',";
					sQry += "'1000',";
					sQry += "'40',";
					sQry += "'101',";
					sQry += "'" + oRecordSet01.Fields.Item("U_BPLId").Value.ToString().Trim() + "',";
					sQry += "'" + oDS_PS_PP077H.GetValue("U_ItemCode", i).ToString().Trim() + "',";
					sQry += "'" + dataHelpClass.Make_ItemName(oDS_PS_PP077H.GetValue("U_ItemName", i).ToString().Trim()) + "',";
					sQry += "'" + oDS_PS_PP077H.GetValue("U_PorNum", i).ToString().Trim() + "',";
					sQry += "'" + oRecordSet01.Fields.Item("U_OrdSub1").Value.ToString().Trim() + "',";
					sQry += "'" + oRecordSet01.Fields.Item("U_OrdSub2").Value.ToString().Trim() + "',";
					sQry += "'" + oRecordSet01.Fields.Item("U_PP030HNo").Value.ToString().Trim() + "',";
					sQry += "'10',";
					sQry += "'N',";
					sQry += "'" + oDS_PS_PP077H.GetValue("U_InDate", i).ToString().Trim() + "'";
					sQry += ")";
					oRecordSet02.DoQuery(sQry);

					// Insert PS_PP040M
					sQry = "INSERT INTO [@PS_PP040M]";
					sQry += " (";
					sQry += " DocEntry,";
					sQry += " LineId,";
					sQry += " VisOrder,";
					sQry += " Object,";
					sQry += " U_LineNum,";
					sQry += " U_LineId,";
					sQry += " U_WorkCode,";
					sQry += " U_WorkName";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += "'" + PP040H_DocEntry + "',";
					sQry += "'1',";
					sQry += "'0',";
					sQry += "'PS_PP040',";
					sQry += "'1',";
					sQry += "'1',";
					sQry += "'" + oForm.Items.Item("EmpId").Specific.Value.ToString().Trim() + "',";
					sQry += "'" + oForm.Items.Item("EmpName").Specific.Value.ToString().Trim() + "'";
					sQry += ")";
					oRecordSet02.DoQuery(sQry);

					//바렐 및 포장 등록
					sQry = "EXEC [PS_PP077_04] '" + oDS_PS_PP077H.GetValue("U_PorNUm", i).ToString().Trim() + "'";
					oRecordSet03.DoQuery(sQry);

					if (oRecordSet03.RecordCount == 0)
					{
						errMessage = "해당작지번호에 대한 작업지시등록 공정중 바렐 또는 포장공정이 없거나 이미등록되어 있습니다. 확인하여 주십시오.";
						throw new Exception();
					}

					j = 0;
					while (!(oRecordSet03.EoF))
					{
						// Insert PS_PP040L
						sQry = "INSERT INTO [@PS_PP040L]";
						sQry += " (";
						sQry += " DocEntry,";
						sQry += " LineId,";
						sQry += " VisOrder,";
						sQry += " Object,";
						sQry += " U_LineNum,";
						sQry += " U_LineId,";
						sQry += " U_OrdMgNum,";
						sQry += " U_Sequence,";
						sQry += " U_CpCode,";
						sQry += " U_CpName,";
						sQry += " U_OrdGbn,";
						sQry += " U_BPLId,";
						sQry += " U_ItemCode,";
						sQry += " U_ItemName,";
						sQry += " U_OrdNum,";
						sQry += " U_OrdSub1,";
						sQry += " U_OrdSub2,";
						sQry += " U_PP030HNo,";
						sQry += " U_PP030MNo,";
						sQry += " U_PQty,";
						sQry += " U_PWeight,";
						sQry += " U_YQty,";
						sQry += " U_YWeight,";
						sQry += " U_NQty,";
						sQry += " U_NWeight";
						sQry += " ) ";
						sQry += "VALUES(";
						sQry += "'" + PP040H_DocEntry + "',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + j + "',";
						sQry += "'PS_PP040',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + oRecordSet01.Fields.Item("U_PP030HNo").Value.ToString().Trim() + "-" + oRecordSet03.Fields.Item("LineId").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("U_Sequence").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("U_CpCode").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("U_CpName").Value.ToString().Trim() + "',";
						sQry += "'101',";
						sQry += "'" + oRecordSet01.Fields.Item("U_BPLId").Value.ToString().Trim() + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_ItemCode", i).ToString().Trim() + "',";
						sQry += "'" + dataHelpClass.Make_ItemName(oDS_PS_PP077H.GetValue("U_ItemName", i).ToString().Trim()) + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_PorNum", i).ToString().Trim() + "',";
						sQry += "'" + oRecordSet01.Fields.Item("U_OrdSub1").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet01.Fields.Item("U_OrdSub2").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet01.Fields.Item("U_PP030HNo").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("LineId").Value.ToString().Trim() + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim() + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim() + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkQty", i).ToString().Trim() + "',";
						sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim() + "',";
						sQry += "'" + 0 + "',";
						sQry += "'" + 0 + "'";
						sQry += ")";
						oRecordSet02.DoQuery(sQry);

						// Insert PS_PP040N
						sQry = "INSERT INTO [@PS_PP040N]";
						sQry += " (";
						sQry += " DocEntry,";
						sQry += " LineId,";
						sQry += " VisOrder,";
						sQry += " Object,";
						sQry += " U_LineNum,";
						sQry += " U_LineId,";
						sQry += " U_OrdMgNum,";
						sQry += " U_CpCode,";
						sQry += " U_CpName";
						sQry += " ) ";
						sQry += "VALUES(";
						sQry += "'" + PP040H_DocEntry + "',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + j + "',";
						sQry += "'PS_PP040',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + j + 1 + "',";
						sQry += "'" + oRecordSet01.Fields.Item("U_PP030HNo").Value.ToString().Trim() + "-" + oRecordSet03.Fields.Item("LineId").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("U_CpCode").Value.ToString().Trim() + "',";
						sQry += "'" + oRecordSet03.Fields.Item("U_CpName").Value.ToString().Trim() + "'";
						sQry += ")";
						oRecordSet02.DoQuery(sQry);

						j += 1;
						oRecordSet03.MoveNext();
					}

					//작업일보-문서번호를 [@PS_PP077H]에 update
					sQry = "Update [@PS_PP077H] Set U_PP040No = '" + PP040H_DocEntry + "' ";
					sQry = sQry + "Where DocNum = '" + PP077HDocNum + "'";
					oRecordSet02.DoQuery(sQry);

					//AutoKey Update
					sQry = "Update [ONNM] Set AutoKey = '" + AutoKey + "' Where ObjectCode = 'PS_PP040'";
					oRecordSet02.DoQuery(sQry);
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet03);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string vReturnValue;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Search")
					{
						PS_PP077_SearchGridData();
					}
					else if (pVal.ItemUID == "Distribute")
					{
						if (!string.IsNullOrEmpty(oForm.Items.Item("DistrQty").Specific.Value.ToString().Trim()) && Convert.ToDouble(oForm.Items.Item("DistrQty").Specific.Value.ToString().Trim()) != 0)
						{
							PS_PP077_DistiributeData();
						}
					}
					else if (pVal.ItemUID == "Save")
					{
						if (PS_PP077_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						if (PS_PP077_DelMatrixSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						if (PS_PP077_SaveData("A", false) == false)
						{
							BubbleEvent = false;
							return;
						}

						vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("이 데이터를 추가한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소"));
						switch (vReturnValue)
						{
							case "1":
								if (PS_PP077_InsertoInventoryGenEntry(2) == false)
								{
									BubbleEvent = false;
									return;
								}
								else
								{
									PSH_Globals.SBO_Application.MessageBox("선택저장 작업이 완료되었습니다.");
									PS_PP077_SearchGridData();
								}
								break;
							case "2":
								BubbleEvent = false;
								break;
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
						if (pVal.ItemUID == "ItemCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
							{
								PS_SM010 ChildForm01 = new PS_SM010();
								ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
							}
						}
						if (pVal.ItemUID == "EmpId")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("EmpId").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
					if (pVal.ItemUID == "DocStatus")
					{
						oForm.Freeze(true);

						if (oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim() == "1")
						{
							oForm.Items.Item("s2").Specific.Caption = "이동등록일";
							oForm.Items.Item("Save").Enabled = true;
						}
						else if (oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim() == "2")
						{
							oForm.Items.Item("s2").Specific.Caption = "포장등록일";
							oForm.Items.Item("Save").Enabled = false;
						}

						if (oGrid.Columns.Count != 0)
						{
							oGrid.DataTable.Clear();
						}

						oMat.Clear();
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
					if (pVal.ItemUID == "Grid01")
					{
						PS_PP077_SearchMatrixData();
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "EmpId" || pVal.ItemUID == "DocDate")
						{
							PS_PP077_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						//라인
						if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "PP070No" || pVal.ColUID == "NPkQty" || pVal.ColUID == "NPkWt"))
						{
							PS_PP077_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP077_ResizeForm(FormUID, ref pVal, ref BubbleEvent);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP077H);
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
			string vReturnValue;

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							if (PS_PP077_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP077_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}

							vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소"));
							switch (vReturnValue)
							{
								case "1":
									if (PS_PP077_InsertoInventoryGenExit(2) == false)
									{
										BubbleEvent = false;
										return;
									}
									else
									{
										PSH_Globals.SBO_Application.MessageBox("선택취소 작업이 완료되었습니다.");
									}
									break;
								case "2":
									BubbleEvent = false;
									break;
							}
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
