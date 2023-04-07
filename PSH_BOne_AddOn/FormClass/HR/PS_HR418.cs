using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 계수조정처리
	/// </summary>
	internal class PS_HR418 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_HR418.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_HR418_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_HR418");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_HR418_CreateItems();
				PS_HR418_ComboBox_Setting();
				PS_HR418_SetDocument(oFromDocEntry01);

				oForm.EnableMenu("1281", false);
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
		/// PS_HR418_CreateItems
		/// </summary>
		private void PS_HR418_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("ZTEMP1");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("ZTEMP1");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//년도
				oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

				//차수
				oForm.DataSources.UserDataSources.Add("Number", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Number").Specific.DataBind.SetBound(true, "", "Number");

				//평가권한
				oForm.DataSources.UserDataSources.Add("Evaluate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Evaluate").Specific.DataBind.SetBound(true, "", "Evaluate");

				//평가그룹
				oForm.DataSources.UserDataSources.Add("Group", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Group").Specific.DataBind.SetBound(true, "", "Group");

				//사번
				oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

				//성명
				oForm.DataSources.UserDataSources.Add("FULLNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("FULLNAME").Specific.DataBind.SetBound(true, "", "FULLNAME");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_HR418_ComboBox_Setting
		/// </summary>
		private void PS_HR418_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//평가차수
				oForm.Items.Item("Number").Specific.ValidValues.Add("1", "1차");
				oForm.Items.Item("Number").Specific.ValidValues.Add("2", "2차");

				//평가권한
				oForm.Items.Item("Evaluate").Specific.ValidValues.Add("1", "1차평가자");
				oForm.Items.Item("Evaluate").Specific.ValidValues.Add("2", "2차평가자");
				oForm.Items.Item("Evaluate").Specific.ValidValues.Add("3", "종합평가자");

				// 평가그룹
				oForm.Items.Item("Group").Specific.ValidValues.Add("1", "1.반장");
				oForm.Items.Item("Group").Specific.ValidValues.Add("2", "2.사원");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_HR418_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_HR418_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_HR418_MTX01
		/// </summary>
		private void PS_HR418_MTX01()
		{
			int Cnt;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string sQry = string.Empty;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				Param01 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();    //사업장
				Param02 = oForm.Items.Item("Year").Specific.Value.ToString().Trim();     //년도
				Param03 = oForm.Items.Item("Number").Specific.Value.ToString().Trim();   //평가차수 1차, 2차평가
				Param04 = oForm.Items.Item("Evaluate").Specific.Value.ToString().Trim(); //평가권한 1차,2차,3차권한
				Param05 = oForm.Items.Item("Group").Specific.Value.ToString().Trim();    //평가그룹 1.반장,2.사원

				if (string.IsNullOrEmpty(Param03) || string.IsNullOrEmpty(Param04) || string.IsNullOrEmpty(Param05))
				{
					errMessage = "평가차수,평가권한,평가그룹을 입력하세요.";
					throw new Exception();
				}

				if (Param04 == "1")
				{
					sQry = " select COUNT(*) ";
					sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
					sQry += " Where a.U_BPLId = '" + Param01 + "'";
					sQry += " and a.U_Year ='" + Param02 + "'";
					sQry += " and a.U_Number = '" + Param03 + "'";
					sQry += " and Isnull(b.U_MSTCOD1,'') <> '' ";
					sQry += " and Isnull(b.U_Complet1,'N') = 'N' ";
					sQry += " and Isnull(b.U_Group, '') = '" + Param05 + "'";
				}
				else if (Param04 == "2")
				{
					sQry = " select COUNT(*) ";
					sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
					sQry += " Where a.U_BPLId = '" + Param01 + "'";
					sQry += " and a.U_Year ='" + Param02 + "'";
					sQry += " and a.U_Number = '" + Param03 + "'";
					sQry += " and Isnull(b.U_MSTCOD2,'') <> '' ";
					sQry += " and Isnull(b.U_Complet2,'N') = 'N' ";
					sQry += " and Isnull(b.U_Group, '') = '" + Param05 + "'";
				}
				else if (Param04 == "3")
				{
					sQry = " select COUNT(*) ";
					sQry += " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code ";
					sQry += " Where a.U_BPLId = '" + Param01 + "'";
					sQry += " and a.U_Year ='" + Param02 + "'";
					sQry += " and a.U_Number = '" + Param03 + "'";
					sQry += " and Isnull(b.U_MSTCOD3,'') <> '' ";
					sQry += " and Isnull(b.U_Complet3,'N') = 'N' ";
					sQry += " and Isnull(b.U_Group, '') = '" + Param05 + "'";
				}
				oRecordSet.DoQuery(sQry);

				Cnt = Convert.ToInt32(oRecordSet.Fields.Item(0).Value);

				if (Cnt > 0)
				{
					oForm.Items.Item("Complete").Specific.VALUE = "평가미완료";
				}
				else
				{
					oForm.Items.Item("Complete").Specific.VALUE = "평가완료";
				}

				sQry = "EXEC PS_HR418_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "'";
				oGrid.DataTable.ExecuteQuery(sQry);
				PS_HR418_GridSetting();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
				oForm.Update();
				PSH_Globals.SBO_Application.StatusBar.SetText("조회를 성공하였습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// 패스워드 Check
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_HR418_PasswordChk(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;
			string sQry;
			string MSTCOD;
			string PassWd;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
				PassWd = oForm.Items.Item("PassWd").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(MSTCOD))
				{
					errMessage = "사번이 없습니다. 입력바랍니다!";
					throw new Exception();
				}

				sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
				sQry += " And  BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "' ";
				sQry += " And  PassWd = '" + oForm.Items.Item("PassWd").Specific.Value.ToString().Trim() + "' ";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value) <= 0)
				{
					functionReturnValue = false;
				}
				else
				{
					functionReturnValue = true;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Grid 꾸며주기
		/// </summary>
		private void PS_HR418_GridSetting()
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
		/// 계수조정 처리
		/// </summary>
		private void PS_HR418_Adjust_Value()
		{
			int i;
			string BPLID;
			string Year;
			string MSTCOD;
			string Number;
			string Evaluate;
			string Group;
			double StaffAvg;
			double RspAvg;
			string TeamCode;
			string RspCode;
			string PeakYN;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
				Evaluate = oForm.Items.Item("Evaluate").Specific.Value.ToString().Trim();
				Group = oForm.Items.Item("Group").Specific.Value.ToString().Trim();

				sQry = "EXEC PS_HR418_02 '" + BPLID + "', '" + Year + "', '" + Number + "', '" + Evaluate + "', '" + Group + "'";
				oRecordSet.DoQuery(sQry);

				for (i = 0; i <= oGrid.Rows.Count - 1; i++)
				{
					MSTCOD = oGrid.DataTable.GetValue(0, i).ToString().Trim();
					StaffAvg = Convert.ToDouble(oGrid.DataTable.GetValue(4, i).ToString().Trim());
					RspAvg = Convert.ToDouble(oGrid.DataTable.GetValue(5, i).ToString().Trim());
					TeamCode = oGrid.DataTable.GetValue(7, i).ToString().Trim();
					RspCode = oGrid.DataTable.GetValue(8, i).ToString().Trim();
					PeakYN = oGrid.DataTable.GetValue(9, i).ToString().Trim();

					switch (Evaluate)
					{
						case "1":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue1 = U_Value1 * U_Rate1,";
							sQry += " U_AAvg1 = Round(U_Value1 * U_Rate1 * 4 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and Isnull([@PS_HR410L].U_MSTCOD1,'') =  (case when '" + MSTCOD + "'" + " = '9999' then Isnull(U_MSTCOD1,'') else '" + MSTCOD + "'" + " end )";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet1 = 'Y'";
							sQry += " and [@PS_HR410L].U_S_TeamCd =  (case when '" + TeamCode + "'" + " = '9999' then U_S_TeamCd else '" + TeamCode + "'" + " end )";
							sQry += " and isnull([@PS_HR410L].U_S_RspCd,'') =  (case when '" + RspCode + "'" + " = '9999' then U_S_RspCd else '" + RspCode + "'" + " end )";
							sQry += " and isnull([@PS_HR410L].U_PeakYN,'N') = (case when '" + PeakYN + "'" + "='N' then 'N' else 'Y' end)";
							break;

						case "2":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue2 = U_Value2 * U_Rate2,";
							sQry += " U_AAvg2 = Round(U_Value2 * U_Rate2 * 3 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and Isnull([@PS_HR410L].U_MSTCOD2,'') =  (case when '" + MSTCOD + "'" + " = '9999' then Isnull(U_MSTCOD2,'') else '" + MSTCOD + "'" + " end )";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet2 = 'Y'";
							sQry += " and [@PS_HR410L].U_S_TeamCd =  (case when '" + TeamCode + "'" + " = '9999' then U_S_TeamCd else '" + TeamCode + "'" + " end )";
							sQry += " and isnull([@PS_HR410L].U_PeakYN,'N') = (case when '" + PeakYN + "'" + "='N' then 'N' else 'Y' end)";
							break;

						case "3":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue3 = U_Value3 * U_Rate3,";
							sQry += " U_AAvg3 = Round(U_Value3 * U_Rate3 * 3 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and Isnull([@PS_HR410L].U_MSTCOD3,'') =  (case when '" + MSTCOD + "'" + " = '9999' then Isnull(U_MSTCOD3,'') else '" + MSTCOD + "'" + " end )";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet3 = 'Y'";
							sQry += " and [@PS_HR410L].U_S_TeamCd =  (case when '" + TeamCode + "'" + " = '9999' then U_S_TeamCd else '" + TeamCode + "'" + " end )";
							sQry += " and isnull([@PS_HR410L].U_S_RspCd,'') =  (case when '" + RspCode + "'" + " = '9999' then U_S_RspCd else '" + RspCode + "'" + " end )";
							sQry += " and isnull([@PS_HR410L].U_PeakYN,'N') = (case when '" + PeakYN + "'" + "='N' then 'N' else 'Y' end)";
							break;
					}
					oRecordSet.DoQuery(sQry);
				}
				PS_HR418_MTX01();
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
		/// PS_HR418_Adjust_Cancel
		/// </summary>
		private void PS_HR418_Adjust_Cancel()
		{
			int i;
			string BPLID;
			string Year;
			string MSTCOD;
			string Number;
			string Evaluate;
			string Group;
			double StaffAvg;
			double RspAvg;
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
				Number = oForm.Items.Item("Number").Specific.Value.ToString().Trim();
				Evaluate = oForm.Items.Item("Evaluate").Specific.Value.ToString().Trim();
				Group = oForm.Items.Item("Group").Specific.Value.ToString().Trim();

				for (i = 0; i <= oGrid.Rows.Count - 1; i++)
				{
					MSTCOD = oGrid.DataTable.GetValue(0, i).ToString().Trim();
					StaffAvg = Convert.ToDouble(oGrid.DataTable.GetValue(4, i).ToString().Trim());
					RspAvg = Convert.ToDouble(oGrid.DataTable.GetValue(5, i).ToString().Trim());

					switch (Evaluate)
					{
						case "1":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue1 = U_Value1 * 0,";
							sQry += " U_AAvg1 = Round(U_Value1 * 0 * 4 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet1 = 'Y'";
							break;

						case "2":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue2 = U_Value2 * 0,";
							sQry += " U_AAvg2 = Round(U_Value2 * 0 * 3 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet2 = 'Y'";
							break;

						case "3":
							sQry = " Update [@PS_HR410L] ";
							sQry += " Set U_AValue3 = U_Value3 * 0,";
							sQry += " U_AAvg3 = Round(U_Value3 * 0 * 3 / 10, 1) ";
							sQry += " from [@PS_HR410H] a";
							sQry += " where a.U_BPLId = '" + BPLID + "'";
							sQry += " and [@PS_HR410L].code = a.code and a.U_Year = '" + Year + "'";
							sQry += " and a.U_Number = '" + Number + "'";
							sQry += " and [@PS_HR410L].U_Group = '" + Group + "'";
							sQry += " and [@PS_HR410L].U_Complet3 = 'Y'";
							break;
					}
					oRecordSet.DoQuery(sQry);
				}
				PS_HR418_MTX01();
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

				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//    break;

				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;

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
				//    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;

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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_HR418_PasswordChk(ref pVal) == false)
							{
								oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								errMessage = "패스워드가 틀렸습니다. 확인바랍니다.";
								throw new Exception();
							}
							else
							{
								PS_HR418_MTX01();
							}
						}
					}
					if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_HR418_PasswordChk(ref pVal) == false)
							{
								oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								errMessage = "패스워드가 틀렸습니다. 확인바랍니다.";
								throw new Exception();
							}
							else
							{
								if (oForm.Items.Item("Complete").Specific.VALUE == "평가완료")
								{
									PS_HR418_Adjust_Value(); //조정처리
								}
								else
								{
									PSH_Globals.SBO_Application.SetStatusBarMessage("평가가 완료되지 않아 조정처리를 할 수 없습니다. 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
								}
							}
						}
					}
					if (pVal.ItemUID == "Btn03") //평가자 그룹핑 삭제
					{
						if (PS_HR418_PasswordChk(ref pVal) == false)
						{
							oForm.Items.Item("PassWd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							errMessage = "패스워드가 틀렸습니다. 확인바랍니다.";
							throw new Exception();
						}
						else
						{
							PS_HR418_Adjust_Cancel();
						}
					}
				}
				else if (pVal.BeforeAction == false)
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
						if (pVal.ItemUID == "MSTCOD")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "TeamCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "RspCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("RspCode").Specific.Value.ToString().Trim()))
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "MSTCOD")
						{
							sQry = "Select FULLNAME = t.U_FullName ";
							sQry += " From [@PH_PY001A] t Where Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "' ";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("FULLNAME").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "TeamCode")
						{
							sQry = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '부서' and b.U_Code = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("TeamNm").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "RspCode")
						{
							sQry = "Select b.U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code Where a.Name = '담당' and b.U_Code = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("RspNm").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						oForm.Update();
					}
				}
				else if (pVal.BeforeAction == false)
				{
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
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
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
