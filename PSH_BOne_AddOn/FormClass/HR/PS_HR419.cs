using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 전문직평가등급처리
	/// </summary>
	internal class PS_HR419 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DBDataSource oDS_PS_HR419L;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR419.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_HR419_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_HR419");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_HR419_CreateItems();
				PS_HR419_ComboBox_Setting();

				oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_HR419_CreateItems
		/// </summary>
		private void PS_HR419_CreateItems()
		{
			try
			{
				oDS_PS_HR419L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("ZTEMP1");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("ZTEMP1");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_HR419_ComboBox_Setting
		/// </summary>
		private void PS_HR419_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("Number").Specific.ValidValues.Add("1", "1차평가");
				oForm.Items.Item("Number").Specific.ValidValues.Add("2", "2차평가");
				oForm.Items.Item("Number").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("Group").Specific.ValidValues.Add("1", "1반장");
				oForm.Items.Item("Group").Specific.ValidValues.Add("2", "2사원");
				oForm.Items.Item("Group").Specific.ValidValues.Add("3", "3피크사원");
				oForm.Items.Item("Group").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_HR419_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_HR419_FlushToItemValue(string oUID)
		{
			string BPLID;
			string Year;
			string Number;
			string MSTCOD;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "MSTCOD":
						BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
						Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
						Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
						MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() +"'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("FULLNAME").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_HR419_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_HR419_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
					oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "평가년도는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
						
				if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                {
					oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "평가자사번은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("EmpNo").Specific.Value.ToString().Trim()))
                {
					errMessage = "피평가자가 선택되지 않았습니다. 확인하여 주십시오.";
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_HR419_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_HR419_MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			int i;
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
						if (string.IsNullOrEmpty(oDS_PS_HR419L.GetValue("U_ColReg01", i).ToString().Trim()))
						{
							errMessage = "평가결과가 올바르지 않습니다. 확인하여 주십시오.";
							throw new Exception();
						}
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_HR419_Search_Grid_Data
		/// </summary>
		private void PS_HR419_Search_Grid_Data()
		{
			int Cnt;
			string BPLID;
			string Year;
			string Number;
			string Group;
			string ProcYN;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				ProcYN = "Y";
				oMat.Clear();
				oForm.Items.Item("TeamCode").Specific.Value = "";
				oForm.Items.Item("TeamName").Specific.Value = "";
				oForm.Items.Item("RspCode").Specific.Value = "";
				oForm.Items.Item("RspName").Specific.Value = "";

				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
				Group = oForm.Items.Item("Group").Specific.Value.ToString().Trim();

				sQry = " select COUNT(*) ";
				sQry += " from [@PS_HR410H] a ";
				sQry += " Where a.U_BPLId = '" + BPLID + "'";
				sQry += " and a.U_Year ='" + Year + "'";
				sQry += " and a.U_Number = '" + Number + "'";
				sQry += " and Isnull(a.U_CloseYN,'') = 'Y'";
				oRecordSet.DoQuery(sQry);

				Cnt = oRecordSet.Fields.Item(0).Value;

				if (Cnt > 0)
				{
					oForm.Items.Item("Complete").Specific.Value = "평가종료";
					ProcYN = "Y";
				}
				else
				{
					oForm.Items.Item("Complete").Specific.Value = "평가진행중";
					ProcYN = "N";
					errMessage = "평가종료후 작업바랍니다.";
					throw new Exception();
				}

				if (ProcYN == "Y")
				{
					sQry = "EXEC PS_HR419_01 '" + BPLID + "','" + Year + "','" + Number + "','" + Group + "'";
					oGrid.DataTable.ExecuteQuery(sQry);
					PS_HR419_GridSetting();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_HR419_GridSetting
		/// </summary>
		private void PS_HR419_GridSetting()
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

					if (sColsTitle == "1차" || sColsTitle == "2차" || sColsTitle == "3차" || sColsTitle == "평균")
					{
						oGrid.Columns.Item(i).RightJustified = true;
					}

					if (oGrid.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
					{
						oGrid.Columns.Item(i).RightJustified = true;
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
		/// PS_HR419_Search_Matrix_Data
		/// </summary>
		private void PS_HR419_Search_Matrix_Data()
		{
			int i;
			int j;
			int Cnt;
			string BPLID;
			string Year;
			string Number;
			string Group;
			string TeamCode;
			string TeamName;
			string RspCode;
			string RspName;
			string PeakYN;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
				Group = oForm.Items.Item("Group").Specific.Value.ToString().Trim();

				for (i = 0; i <= oGrid.Rows.Count - 1; i++)
				{
					if (oGrid.Rows.IsSelected(i) == true)
					{
						TeamCode = oGrid.DataTable.GetValue(0, i).ToString().Trim();
						TeamName = oGrid.DataTable.GetValue(1, i).ToString().Trim();
						RspCode = oGrid.DataTable.GetValue(2, i).ToString().Trim();
						RspName = oGrid.DataTable.GetValue(3, i).ToString().Trim();
						Cnt = Convert.ToInt32(oGrid.DataTable.GetValue(4, i).ToString().Trim());
						PeakYN = oGrid.DataTable.GetValue(6, i).ToString().Trim();

						oForm.Items.Item("TeamCode").Specific.Value = TeamCode;
						oForm.Items.Item("TeamName").Specific.Value = TeamName;
						oForm.Items.Item("RspCode").Specific.Value = RspCode;
						oForm.Items.Item("RspName").Specific.Value = RspName;
						oForm.Items.Item("Cnt").Specific.Value = Cnt;

						sQry = "EXEC PS_HR419_02 '" + BPLID + "', '" + Year + "', '" + Number + "', '" + Group + "', '" + TeamCode + "', '" + RspCode + "','" + PeakYN + "'";
						oRecordSet.DoQuery(sQry);
						Cnt = oDS_PS_HR419L.Size;

						if (Cnt > 0)
						{
							for (j = 0; j <= Cnt - 1; j++)
							{
								oDS_PS_HR419L.RemoveRecord(oDS_PS_HR419L.Size - 1);
							}
							if (Cnt == 1)
							{
								oDS_PS_HR419L.Clear();
							}
						}

						oMat.LoadFromDataSource();

						j = 1;
						while (!oRecordSet.EoF)
						{
							if (oDS_PS_HR419L.Size < j)
							{
								oDS_PS_HR419L.InsertRecord(j - 1);
							}
							oDS_PS_HR419L.SetValue("U_LineNum", j - 1, Convert.ToString(j));
							oDS_PS_HR419L.SetValue("U_ColReg01", j - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_HR419L.SetValue("U_ColReg02", j - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_HR419L.SetValue("U_ColReg03", j - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
							oDS_PS_HR419L.SetValue("U_ColQty01", j - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
							oDS_PS_HR419L.SetValue("U_ColReg04", j - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
							j += 1;
							oRecordSet.MoveNext();
						}

						oMat.LoadFromDataSource();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_HR419_Save_Data
		/// </summary>
		/// <returns></returns>
		private bool PS_HR419_Save_Data()
		{
			bool functionReturnValue = false;
			string sQry;
			int i;
			string BPLID;
			string Year;
			string Number;
			string MSTCOD;
			string Grade;
			decimal Ranking;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();

				for (i = 1; i <= oMat.VisualRowCount; i++)
				{
					MSTCOD = oMat.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value.ToString().Trim();
					Ranking = Convert.ToDecimal(oMat.Columns.Item("Ranking").Cells.Item(i).Specific.Value);
					Grade = oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value.ToString().Trim();

					sQry = " Update [@PS_HR410L] Set U_Ranking =  " + Ranking + ",";
					sQry += " U_Grade = '" + Grade + "'";
					sQry += " From [@PS_HR410H] a ";
					sQry += " Where a.Code = [@PS_HR410L].Code ";
					sQry += " And a.U_BPLId = '" + BPLID + "'";
					sQry += " And a.U_Year = '" + Year + "'";
					sQry += " And a.U_Number = '" + Number + "'";
					sQry += " And [@PS_HR410L].U_MSTCOD = '" + MSTCOD + "'";
					oRecordSet.DoQuery(sQry);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("정상처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_HR419_Matrix_Grade_Set
		/// </summary>
		private void PS_HR419_Matrix_Grade_Set()
		{
			string sQry;
			int i;
			int Cnt;
			string BPLID;
			int A;
			int B;
			int C;
			int D;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Cnt = Convert.ToInt32(oForm.Items.Item("Cnt").Specific.Value.ToString().Trim());

				sQry = "Select b.U_A, b.U_B, b.U_C, b.U_D";
				sQry += " From [@PS_HR402H] a inner Join [@PS_HR402L] b On a.Code = b.Code";
				sQry += " Where a.U_BPLId = '" + BPLID + "'";
				sQry += " And b.U_Number = " + Cnt;
				oRecordSet.DoQuery(sQry);

				A = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);
				B = Convert.ToInt32(oRecordSet.Fields.Item(1).Value);
				C = Convert.ToInt32(oRecordSet.Fields.Item(2).Value);
				D = Convert.ToInt32(oRecordSet.Fields.Item(3).Value);

				for (i = 1; i <= A + B + C + D; i++)
				{
					oMat.Columns.Item("Ranking").Cells.Item(i).Specific.Value = i;
					if (i <= A)
					{
						oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "A";
					}
					else if (i <= A + B)
					{
						oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "B";
					}
					else if (i <= A + B + C)
					{
						oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "C";
					}
					else if (i <= A + B + C + D)
					{
						oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "D";
					}
					else
					{
						oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "Z";
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_HR419_TValue_Set
		/// </summary>
		private void PS_HR419_TValue_Set()
		{
			string sQry;
			int i;
			string BPLID;
			string Year;
			string Number;
			string MSTCOD;
			string Grade;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();

				for (i = 1; i <= oMat.VisualRowCount; i++)
				{
					MSTCOD = oMat.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value.ToString().Trim();
					Grade = oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value.ToString().Trim();

					sQry = " Update [@PS_HR410L] Set U_TValue = U_AAvg1 + U_AAvg2 + U_AAvg3 + U_Fix + U_Adjust";
					sQry += " From [@PS_HR410H] a ";
					sQry += " Where a.Code = [@PS_HR410L].Code ";
					sQry += " And a.U_BPLId = '" + BPLID + "'";
					sQry += " And a.U_Year = '" + Year + "'";
					sQry += " And a.U_Number = '" + Number + "'";
					sQry += " And [@PS_HR410L].U_MSTCOD = '" + MSTCOD + "'";
					oRecordSet.DoQuery(sQry);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("정상처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;

				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;

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

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

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
			string Complete;

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01") //조회버튼
					{
						PS_HR419_Search_Grid_Data();
					}
					else if (pVal.ItemUID == "Btn02") //저장버튼
					{
						Complete = oForm.Items.Item("Complete").Specific.Value.ToString().Trim();

						if (Complete != "평가종료")
						{
							PS_HR419_Search_Grid_Data(); //새로고침
							PSH_Globals.SBO_Application.SetStatusBarMessage("평가진행중입니다. 평가완료후 처리바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
						}
						else
						{
							if (PS_HR419_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_HR419_Save_Data() == false)
							{
								BubbleEvent = false;
								return;
							}
							else
							{
								PS_HR419_Search_Grid_Data(); //새로고침
							}
						}
					}
					else if (pVal.ItemUID == "Btn03") //평가등급 RESET
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("Ranking").Cells.Item(i).Specific.Value = 0;
							oMat.Columns.Item("Grade").Cells.Item(i).Specific.Value = "";
						}
						oForm.Freeze(false);
					}
					else if (pVal.ItemUID == "Btn04") //완료처리
					{
						PS_HR419_Matrix_Grade_Set();
					}
					else if (pVal.ItemUID == "Btn05") //평가총점계산
					{
						PS_HR419_TValue_Set();
						PS_HR419_Search_Matrix_Data();
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
						if (pVal.ItemUID == "RateCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("RateCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Grid01")
					{
						PS_HR419_Search_Matrix_Data();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
						if (pVal.ItemUID == "Year")
						{
							if (!string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
							{
								sQry = "select U_Number from [@PS_HR410H] a";
								sQry += " Where Isnull(a.U_OpenYN,'N') = 'Y' and isnull(a.U_CloseYN,'N') = 'N' ";
								sQry += " and a.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "' ";
								sQry += " and a.U_Year = '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "' ";
								oRecordSet.DoQuery(sQry);

								if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
								{
									oForm.Items.Item("Number").Specific.Select(oRecordSet.Fields.Item(0).Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
								}
							}
						}

						if (pVal.ItemUID == "MSTCOD")
						{
							PS_HR419_FlushToItemValue(pVal.ItemUID);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
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
					oForm.Items.Item("Grid01").Top = 79;
					oForm.Items.Item("Grid01").Left = 10;
					oForm.Items.Item("Grid01").Width = oForm.Width / 2;

					oForm.Items.Item("Mat01").Top = 79;
					oForm.Items.Item("Mat01").Left = (oForm.Width / 2) + 30;
					oForm.Items.Item("Mat01").Width = (oForm.Width / 2) - 50;

					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_HR419L);
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
			string vReturnValue;

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							if (PS_HR419_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_HR419_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소"));
							if (vReturnValue == "2")
							{
								BubbleEvent = false;
								return;
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
		}
	}
}
