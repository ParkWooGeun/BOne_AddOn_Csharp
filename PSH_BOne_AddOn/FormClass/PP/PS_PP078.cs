using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 휘팅서울포장등록취소
	/// </summary>
	internal class PS_PP078 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.Grid oGrid;
			
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP078.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP078_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP078");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();

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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			try
			{
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("1", "포장대기");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("2", "포장완료");
				oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
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
		private void FlushToItemValue(string oUID, int oRow, string oCol)
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
						break;
					case "EmpId":
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("EmpId").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("EmpName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool MatrixSpaceLineDel()
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
					// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						if (oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() == "Y")
						{
							oDS_PS_PP077H.Offset = i;
							j += 1;	//체크된 라인Count
						}
					}
				}

				// 체크된 라인Count
				if (j == 0)
				{
					errMessage = "선택되어진 라인이 없습니다. 확인하여 주십시오";
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
		/// Form_Resize
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Form_Resize(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Items.Item("Grid01").Top = 40;
				oForm.Items.Item("Grid01").Height = (oForm.Height / 2) - 70;
				oForm.Items.Item("Grid01").Left = 6;
				oForm.Items.Item("Grid01").Width = oForm.Width - 21;
				oForm.Items.Item("DistrQty").Top = (oForm.Height / 2) - 28;
				oForm.Items.Item("Distribute").Top = (oForm.Height / 2) - 33;
				oForm.Items.Item("Mat01").Top = (oForm.Height / 2) - 10;
				oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 60;
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
		}

		/// <summary>
		/// Search_Grid_Data
		/// </summary>
		private void Search_Grid_Data()
		{
			string sQry;
			string MovDocNo;
			string FrDate;
			string ToDate;
			string ItemCode;
			string Status;

			try
			{
				oForm.Freeze(true);

				oForm.Items.Item("DistrQty").Specific.Value = "";
				oForm.Items.Item("SumPkQty").Specific.Value = "";
				oForm.Items.Item("SumPkWt").Specific.Value = "";
				oForm.Items.Item("EmpId").Specific.Value = "";
				oForm.Items.Item("EmpName").Specific.Value = "";

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

				sQry = "EXEC PS_PP078_01 '" + MovDocNo + "','" + FrDate + "','" + ToDate + "','" + ItemCode + "', '" + Status + "'";

				// Procedure 실행(Grid 사용)
				oForm.DataSources.DataTables.Item(0).ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("ZTEMP");

				GridSetting();
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
		/// GridSetting
		/// </summary>
		private void GridSetting()
		{
			int i;
			string sColsTitle;

			try
			{
				oForm.Freeze(true);

				oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

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
		/// Search_Matrix_Data
		/// </summary>
		private void Search_Matrix_Data()
		{
			string sQry;

			int i;
			int j;
			int Cnt;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				for (i = 0; i <= oGrid.Rows.Count - 1; i++)
				{
					if (oGrid.Rows.IsSelected(i) == true)
					{
						sQry = "EXEC PS_PP078_02 '" + oGrid.DataTable.GetValue(0, i).ToString().Trim() + "', '" + oGrid.DataTable.GetValue(1, i).ToString().Trim() + "', '" + oGrid.DataTable.GetValue(2, i).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						// Line 초기화
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
						//Matrix에 Data 뿌려준다
						j = 1;
						while (!(oRecordSet.EoF))
						{
							if (oDS_PS_PP077H.Size < j)
							{
								oDS_PS_PP077H.InsertRecord(j - 1); //라인추가
							}
							oDS_PS_PP077H.SetValue("DocEntry", j - 1, Convert.ToString(j));
							oDS_PS_PP077H.SetValue("U_Check", j - 1, "Y");
							oDS_PS_PP077H.SetValue("Canceled", j - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_MovDocNo", j - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PP070No", j - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PP070NoL", j - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PorNum", j - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_ItemCode", j - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_ItemName", j - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PkQty", j - 1, oRecordSet.Fields.Item(7).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_PkWt", j - 1, oRecordSet.Fields.Item(8).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_NPkQty", j - 1, oRecordSet.Fields.Item(9).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_NPkWt", j - 1, oRecordSet.Fields.Item(10).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("U_InDate", j - 1, Convert.ToDateTime(oRecordSet.Fields.Item(13).Value.ToString().Trim()).ToString("yyyyMMdd"));
							oDS_PS_PP077H.SetValue("U_PP040No", j - 1, oRecordSet.Fields.Item(14).Value.ToString().Trim());
							oDS_PS_PP077H.SetValue("DocNum", j - 1, oRecordSet.Fields.Item(15).Value.ToString().Trim());

							j += 1;
							oRecordSet.MoveNext();
						}
						oMat.LoadFromDataSource();
					}
				}
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
		/// Insert_oInventoryGenExit
		/// </summary>
		/// <param name="ChkType"></param>
		/// <returns></returns>
		private bool Insert_oInventoryGenExit(int ChkType)
		{
			bool functionReturnValue = false;
			//   출고 DI
			SAPbobsCOM.Documents DI_oInventoryGenExit = null; //재고입고 문서 객체
			int RetVal;
			string sQry;
			int errCode;
			string errMsg = string.Empty;
			int i;
			short oRow;
			string oDate;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//출고
				PSH_Globals.oCompany.StartTransaction();
				DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); // 문서타입(입고)

				i = 1;
				var _with1 = DI_oInventoryGenExit;

				// Header
				_with1.Comments = "포장처리등록취소 출고";
				// Line - 라인별 품목등록시 입고문서 번호가 달라질수 있으므로(분할시) 라인별 출고문서 생성
				for (oRow = 0; oRow <= oDS_PS_PP077H.Size - 1; oRow++)
				{
					if (oDS_PS_PP077H.GetValue("U_Check", oRow).ToString().Trim() == "Y")
					{
						//_with1.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oDS_PS_PP077H.GetValue("U_InDate", oRow)), "0000-00-00"));
						//_with1.TaxDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Trim(oDS_PS_PP077H.GetValue("U_InDate", oRow)), "0000-00-00"));

						oDate = oDS_PS_PP077H.GetValue("U_InDate", oRow).ToString().Trim();
						oDate = Convert.ToDateTime(oDate.Substring(0, 4) + "-" + oDate.Substring(4, 2) + "-" + oDate.Substring(6, 2)).ToString("yyyy-MM-dd");
						_with1.DocDate = Convert.ToDateTime(oDate);
						_with1.TaxDate = Convert.ToDateTime(oDate);

						_with1.Lines.SetCurrentLine(i - 1);
						_with1.Lines.ItemCode = oDS_PS_PP077H.GetValue("U_ItemCode", oRow).ToString().Trim();
						_with1.Lines.ItemDescription = oDS_PS_PP077H.GetValue("U_ItemName", oRow).ToString().Trim();
						_with1.Lines.Quantity = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", oRow).ToString().Trim());
						_with1.Lines.WarehouseCode = "104";	//제품-서울

						//포장등록시 입고 문서번호 select
						sQry = "select U_OIGNNo from [@PS_PP077H] where DocNum = '" + oDS_PS_PP077H.GetValue("DocNum", oRow).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						_with1.UserFields.Fields.Item("U_CancDoc").Value = oRecordSet.Fields.Item("U_OIGNNo").Value.ToString().Trim();

						RetVal = DI_oInventoryGenExit.Add();
						if (0 != RetVal)
						{
							PSH_Globals.oCompany.GetLastError(out errCode, out errMsg);
							throw new Exception();
						}
						if (ChkType != 2)
						{
							PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
						}
						else
						{
							PSH_Globals.oCompany.GetNewObjectCode(out VIDocNum);

							//[@PS_PP077H]에 -수량으로 Insert
							if (Save_Data("B", true, oRow) == false)
							{
								throw new Exception();
							}
							//[@PS_PP040H](작업지시)에 Update!!!
							if (Update_PS_PP040(oRow) == false)
							{
								throw new Exception();
							}
						}
					}
				}

				PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMsg != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMsg);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Save_Data
		/// </summary>
		/// <param name="A_B"></param>
		/// <param name="Cancel"></param>
		/// <param name="PvalRow"></param>
		/// <returns></returns>
		private bool Save_Data(string A_B, bool Cancel, int PvalRow)
		{
			bool functionReturnValue = false;

			string sQry;
			int i;
			int DocNum;
			double PkWt;
			double NPkWt;
			double OPkWt;
			string errMessage = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
									//이동요청중량 보다 기포장중량+포장중량이 클수없음, Check
									PkWt = Convert.ToDouble( oDS_PS_PP077H.GetValue("U_PkWt", i).ToString().Trim());
									NPkWt = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", i).ToString().Trim());
									OPkWt = Convert.ToDouble(oDS_PS_PP077H.GetValue("U_OPkWt", i).ToString().Trim());

									if (PkWt < NPkWt + OPkWt)
									{
										errMessage = "기포장중량 + 포장중량은 이동요청중량을 초과할수 없습니다.. 확인하여 주십시오.";
										throw new Exception();
									}
									break;

								case "B":
									//[@PS_PP077H]에 -수량 & -중량으로 Insert
									//DocNum 생성
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
									sQry += "'" + oDS_PS_PP077H.GetValue("U_Check", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_MovDocNo", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PP070No", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PP070NoL", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PorNum", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_ItemCode", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_ItemName", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PkQty", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PkWt", PvalRow).ToString().Trim() + "',";

									//취소 마이너스(-) insert
									if (Cancel == true)
									{
										sQry += "'" + -1 * Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkQty", PvalRow).ToString().Trim()) + "',";
										sQry += "'" + -1 * Convert.ToDouble(oDS_PS_PP077H.GetValue("U_NPkWt", PvalRow).ToString().Trim()) + "',";
									}
									else
									{
										sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkQty", PvalRow).ToString().Trim() + "',";
										sQry += "'" + oDS_PS_PP077H.GetValue("U_NPkWt", PvalRow).ToString().Trim() + "',";
									}
									sQry += "'" + oDS_PS_PP077H.GetValue("U_OPkQty", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_OPkWt", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_InDate", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("U_PP040No", PvalRow).ToString().Trim() + "',";
									sQry += "'" + oDS_PS_PP077H.GetValue("DocNum", PvalRow).ToString().Trim() + "',";
									sQry += "Null,";
									sQry += "'" + VIDocNum + "')";
									oRecordSet.DoQuery(sQry);

									//Canceled, Status 필드 Update
									//기존문서 update
									sQry = "Update [@PS_PP077H] set Canceled='Y', Status='C' where DocNum = '" + oDS_PS_PP077H.GetValue("DocNum", PvalRow).ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);
									//현재insert 되어진 문서 Update
									sQry = "Update [@PS_PP077H] set Canceled='Y', Status='C' where DocNum = '" + DocNum + "'";
									oRecordSet.DoQuery(sQry);

									// 1회 Lopping 후 exit
									functionReturnValue = true;
									return functionReturnValue;
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
		/// Update_PS_PP040
		/// </summary>
		/// <param name="PvalRow"></param>
		/// <returns></returns>
		private bool Update_PS_PP040(int PvalRow)
		{
			bool functionReturnValue = false;

			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//기존 생성된 작업일보 취소
				for (i = 0; i <= oMat.RowCount - 1; i++)
				{
					if (oDS_PS_PP077H.GetValue("U_Check", i).ToString().Trim() == "Y")
					{
						sQry = "Update [@PS_PP040H] set Canceled='Y', Status='C' where DocNum = '" + oDS_PS_PP077H.GetValue("U_PP040No", PvalRow).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						//1회 Lopping 후 exit
						functionReturnValue = true;
						return functionReturnValue;
					}
				}
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
						Search_Grid_Data();
					}
					else if (pVal.ItemUID == "Save")
					{
						if (MatrixSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}

						vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소"));
						switch (vReturnValue)
						{
							case "1":
								if (Insert_oInventoryGenExit(2) == false)
								{
									BubbleEvent = false;
									return;
								}
								else
								{
									PSH_Globals.SBO_Application.MessageBox("선택저장 작업이 완료되었습니다.");
									Search_Grid_Data();
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
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "DocStatus")
					{
						
						if (oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim() == "1")
						{
							oForm.Items.Item("s2").Specific.Caption = "이동등록일";
						}
						else if (oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim() == "2")
						{
							oForm.Items.Item("s2").Specific.Caption = "포장등록일";
						}

						if (oGrid.Columns.Count != 0)
							oGrid.DataTable.Clear();
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
						Search_Matrix_Data();
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
						if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "EmpId")
						{
							FlushToItemValue(pVal.ItemUID, 0, "");
						}
						//라인
						if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "PP070No"))
						{
							FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					Form_Resize(FormUID, ref pVal, ref BubbleEvent);
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
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
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
	}
}
