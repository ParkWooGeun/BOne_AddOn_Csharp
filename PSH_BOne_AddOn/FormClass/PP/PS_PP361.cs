using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작지별 생산진행현황
	/// </summary>
	internal class PS_PP361 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP361.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP361_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP361");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP361_CreateItems();
				PS_PP361_SetComboBox();

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
		/// PS_PP361_CreateItems
		/// </summary>
		private void PS_PP361_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

				oForm.DataSources.DataTables.Add("ZTEMP01");
				oForm.DataSources.DataTables.Add("ZTEMP02");
				oForm.DataSources.DataTables.Add("ZTEMP03");

				oForm.DataSources.UserDataSources.Add("SYYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("SYYYYMM").Specific.DataBind.SetBound(true, "", "SYYYYMM");
				oForm.DataSources.UserDataSources.Item("SYYYYMM").Value = DateTime.Now.ToString("yyyy-MM");

				oForm.DataSources.UserDataSources.Add("EYYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("EYYYYMM").Specific.DataBind.SetBound(true, "", "EYYYYMM");
				oForm.DataSources.UserDataSources.Item("EYYYYMM").Value = DateTime.Now.ToString("yyyy-MM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP361_SetComboBox
		/// </summary>
		private void PS_PP361_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//거래처구분(사업장이 아님)
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("ALL", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//품목구분
				sQry = "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code";
				oRecordSet.DoQuery(sQry);

				oForm.Items.Item("ItemGB").Specific.ValidValues.Add("A", "전체");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("ItemGB").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("ItemGB").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//구분
				oForm.Items.Item("Section").Specific.ValidValues.Add("A", "전체");
				oForm.Items.Item("Section").Specific.ValidValues.Add("B", "미완료");
				oForm.Items.Item("Section").Specific.ValidValues.Add("C", "완료");
				oForm.Items.Item("Section").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//선택
				oForm.Items.Item("Selection").Specific.ValidValues.Add("000", "선택");
				oForm.Items.Item("Selection").Specific.ValidValues.Add("100", "자재비내역");
				oForm.Items.Item("Selection").Specific.ValidValues.Add("200", "자체가공비내역");
				oForm.Items.Item("Selection").Specific.ValidValues.Add("300", "설계비내역");
				oForm.Items.Item("Selection").Specific.ValidValues.Add("400", "외주가공비내역");
				oForm.Items.Item("Selection").Specific.ValidValues.Add("500", "외주제작비내역");
				oForm.Items.Item("Selection").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("SYYYYMM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// PS_PP361_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP361_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("SYYYYMM").Specific.Value.ToString().Trim()))
				{
					errMessage = "해당년월의 시작은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (oForm.Items.Item("SYYYYMM").Specific.Value.ToString().Trim().Length != 7)
				{
					errMessage = "시작년월(YYYY-MM)의 형식을 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("EYYYYMM").Specific.Value.ToString().Trim()))
				{
					errMessage = "해당년월의 종료는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (oForm.Items.Item("EYYYYMM").Specific.Value.ToString().Trim().Length != 7)
				{
					errMessage = "종료년월(YYYY-MM)의 형식을 확인하여 주십시오.";
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
		/// PS_PP361_SetGrid
		/// </summary>
		/// <param name="GridNo"></param>
		private void PS_PP361_SetGrid(string GridNo)
		{
			int i;

			try
			{
				oForm.Freeze(true);

				switch (GridNo)
				{
					case "Grid01":
						((SAPbouiCOM.EditTextColumn)oGrid01.Columns.Item(2)).LinkedObjectType = "4"; // Link to ItemMaster
						for (i = 0; i <= oGrid01.Columns.Count - 1; i++)
						{
							if (oGrid01.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
							{
								oGrid01.Columns.Item(i).RightJustified = true;
							}
						}
						break;

					case "Grid02":
						for (i = 0; i <= oGrid02.Columns.Count - 1; i++)
						{
							if (oGrid02.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
							{
								oGrid02.Columns.Item(i).RightJustified = true;
							}
						}
						break;
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
		/// PS_PP361_ResizeForm
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void PS_PP361_ResizeForm(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				//Grid01
				oForm.Items.Item("Grid01").Top = 64;
				oForm.Items.Item("Grid01").Height = (oForm.Height / 2) - 130;
				oForm.Items.Item("Grid01").Left = 10;
				oForm.Items.Item("Grid01").Width = oForm.Width - 35;

				//Btn01
				oForm.Items.Item("Btn01").Top = oForm.Items.Item("Grid01").Top + oForm.Items.Item("Grid01").Height + 10;
				oForm.Items.Item("Btn01").Left = oForm.Width - 90;

				//Sub 작번
				oForm.Items.Item("S200").Top = oForm.Items.Item("Grid01").Top + oForm.Items.Item("Grid01").Height + 35;

				//Grid02
				oForm.Items.Item("Grid02").Top = oForm.Items.Item("Grid01").Top + oForm.Items.Item("Grid01").Height + 50;
				oForm.Items.Item("Grid02").Height = (oForm.Height / 2) - 300;
				oForm.Items.Item("Grid02").Left = 10;
				oForm.Items.Item("Grid02").Width = oForm.Width - 35;

				//Btn02
				oForm.Items.Item("Btn02").Top = oForm.Items.Item("Grid02").Top + oForm.Items.Item("Grid02").Height + 10;
				oForm.Items.Item("Btn02").Left = oForm.Width - 90;

				//선택
				oForm.Items.Item("22").Top = oForm.Items.Item("Grid02").Top + oForm.Items.Item("Grid02").Height + 35;
				//Selection
				oForm.Items.Item("Selection").Top = oForm.Items.Item("Grid02").Top + oForm.Items.Item("Grid02").Height + 35;

				//Grid03
				oForm.Items.Item("Grid03").Top = oForm.Items.Item("Grid02").Top + oForm.Items.Item("Grid02").Height + 50;
				oForm.Items.Item("Grid03").Height = (oForm.Height / 2) - 280;
				oForm.Items.Item("Grid03").Left = 10;
				oForm.Items.Item("Grid03").Width = oForm.Width - 35;

				//Btn03
				oForm.Items.Item("Btn03").Top = oForm.Items.Item("Grid03").Top + oForm.Items.Item("Grid03").Height + 10;
				oForm.Items.Item("Btn03").Left = oForm.Width - 90;

				if (oGrid01.Rows.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}
				if (oGrid02.Rows.Count > 0)
				{
					oGrid02.AutoResizeColumns();
				}
				if (oGrid03.Rows.Count > 0)
				{
					oGrid03.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP361_SearchGrid01Data
		/// </summary>
		private void PS_PP361_SearchGrid01Data()
		{
			string sQry;
			string SYYYYMM;  //작번등록년월 시작
			string EYYYYMM;  //작번등록년월 종료
			string BPLID;    //사업장
			string ItemGB;   //품목구분
			string Section;  //구분
			string ItemName; //품명
			string Size;     //규격
			string OrdNum;   //작번
			string CardCode; //거래처코드

			try
			{
				oForm.Freeze(true);

				SYYYYMM = oForm.Items.Item("SYYYYMM").Specific.Value.ToString().Trim();
				EYYYYMM = oForm.Items.Item("EYYYYMM").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				ItemGB = oForm.Items.Item("ItemGB").Specific.Selected.Value.ToString().Trim();
				Section = oForm.Items.Item("Section").Specific.Selected.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				Size = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					CardCode = "%";
				}
				else
				{
					CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				}
				if (ItemGB == "A")
				{
					ItemGB = "%";
				}
				if (Section == "A")
				{
					Section = "%";
				}
				if (BPLID == "ALL")
				{
					BPLID = "%";
				}
				if (string.IsNullOrEmpty(ItemName))
				{
					ItemName = "%";
				}
				if (string.IsNullOrEmpty(Size))
				{
					Size = "%";
				}
				if (string.IsNullOrEmpty(OrdNum))
				{
					OrdNum = "%";
				}

				sQry = "EXEC PS_PP361_01 '" + SYYYYMM + "','" + EYYYYMM + "','" + BPLID + "','" + ItemGB + "', '" + Section + "','" + ItemName + "', '" + Size + "', '" + OrdNum + "','" + CardCode + "'";

				oForm.DataSources.DataTables.Item("ZTEMP01").ExecuteQuery(sQry);
				oGrid01.DataTable = oForm.DataSources.DataTables.Item("ZTEMP01");

				PS_PP361_SetGrid("Grid01");
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
		/// PS_PP361_SearchGrid02Data
		/// </summary>
		private void PS_PP361_SearchGrid02Data()
		{
			string sQry;
			int i;

			try
			{
				oForm.Freeze(true);

				for (i = 0; i <= oGrid01.Rows.Count - 1; i++)
				{
					if (oGrid01.Rows.IsSelected(i) == true)
					{
						sQry = "EXEC PS_PP361_02 '" + oGrid01.DataTable.GetValue(0, i).ToString().Trim() + "'";
						oForm.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery(sQry);
						oGrid02.DataTable = oForm.DataSources.DataTables.Item("ZTEMP02");
					}
				}
				PS_PP361_SetGrid("Grid02");
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
		/// PS_PP361_SearchGrid03Data
		/// </summary>
		private void PS_PP361_SearchGrid03Data()
		{
			string sQry = string.Empty;
			int i;
			int j;
			string errMessage = string.Empty;

			try
			{
				oForm.Freeze(true);

				j = 0;
				for (i = 0; i <= oGrid02.Rows.Count - 1; i++)
				{
					if (oGrid02.Rows.IsSelected(i) == true)
					{
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "100")  //자재비내역
						{
							sQry = "EXEC PS_PP361_03 '" + oGrid02.DataTable.GetValue(0, i).ToString().Trim() + "','" + oGrid02.DataTable.GetValue(1, i).ToString().Trim() + "'";
						}
						else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "200")  //자체가공비내역
						{
							sQry = "EXEC PS_PP361_05 '" + oGrid02.DataTable.GetValue(0, i).ToString().Trim() + "','" + oGrid02.DataTable.GetValue(1, i).ToString().Trim() + "'";
						}
						else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "300") //설계비내역
						{
							sQry = "EXEC PS_PP361_04 '" + oGrid02.DataTable.GetValue(0, i).ToString().Trim() + "','" + oGrid02.DataTable.GetValue(1, i).ToString().Trim() + "'";
						}
						else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "400") //외주가공비
						{
							sQry = "EXEC PS_PP361_06 '" + oGrid02.DataTable.GetValue(0, i).ToString().Trim() + "','" + oGrid02.DataTable.GetValue(1, i).ToString().Trim() + "'";
						}
						else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "500") //외주제작비
						{
							sQry = "EXEC PS_PP361_07 '" + oGrid02.DataTable.GetValue(0, i).ToString().Trim() + "','" + oGrid02.DataTable.GetValue(1, i).ToString().Trim() + "'";
						}

						oForm.DataSources.DataTables.Item("ZTEMP03").ExecuteQuery(sQry);
						oGrid03.DataTable = oForm.DataSources.DataTables.Item("ZTEMP03");

						j += 1;
					}
				}
			
				if (j == 0)
				{
					oGrid03.DataTable.Clear();
					errMessage = "[Sub 작번]의 행을 선택한 뒤 [선택] 중 하나의 내역을 선택하여 주십시오.";
					throw new Exception();
				}

				PS_PP361_SetGrid("Grid03");
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
			}
		}

		/// <summary>
		/// PS_PP361_PrintQuery1
		/// </summary>
		[STAThread]
		private void PS_PP361_PrintQuery1()
		{
			string WinTitle;
			string ReportName;

			string SYYYYMM;
			string EYYYYMM;
			string BPLID;
			string ItemGB;
			string Section;
			string ItemName;
			string Size;
			string OrdNum;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				SYYYYMM = oForm.Items.Item("SYYYYMM").Specific.Value.ToString().Trim();
				EYYYYMM = oForm.Items.Item("EYYYYMM").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				ItemGB = oForm.Items.Item("ItemGB").Specific.Value.ToString().Trim();
				Section = oForm.Items.Item("Section").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				Size = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();

				if (ItemGB == "A")
                {
					ItemGB = "%";
				}
				if (Section == "A")
                {
					Section = "%";
				}
				if (BPLID == "ALL")
                {
					BPLID = "%";
				}
				if (string.IsNullOrEmpty(ItemName))
                {
					ItemName = "%";
				}
				if (string.IsNullOrEmpty(Size))
                {
					Size = "%";
				}
				if (string.IsNullOrEmpty(OrdNum))
                {
					OrdNum = "%";
				}

				WinTitle = "[PS_PP361_01] 작번 List";
				ReportName = "PS_PP361_01.RPT";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@SYYYYMM", SYYYYMM));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@SYYYYMM", SYYYYMM));
				dataPackParameter.Add(new PSH_DataPackClass("@EYYYYMM", EYYYYMM));
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemGB", ItemGB));
				dataPackParameter.Add(new PSH_DataPackClass("@Section", Section));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemName", ItemName));
				dataPackParameter.Add(new PSH_DataPackClass("@Size", Size));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP361_PrintQuery2
		/// </summary>
		[STAThread]
		private void PS_PP361_PrintQuery2()
		{
			int i;
			string WinTitle;
			string ReportName;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				WinTitle = "[PS_PP361_02] SUB작번 List";
				ReportName = "PS_PP361_02.RPT";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				for (i = 0; i <= oGrid01.Rows.Count - 1; i++)
				{
					if (oGrid01.Rows.IsSelected(i) == true)
					{
						dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", oGrid01.DataTable.GetValue(0, i).ToString().Trim()));
					}
				}
				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP361_PrintQuery3
		/// </summary>
		[STAThread]
		private void PS_PP361_PrintQuery3()
		{
			int i = 0;
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			try
			{
				if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "100")
				{
					WinTitle = "[PS_PP361_03] 자재비 내역";
					ReportName = "PS_PP361_03.RPT";
				}
				else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "200")
				{
					WinTitle = "[PS_PP361_05] 자체가공비 내역";
					ReportName = "PS_PP361_05.RPT";
				}
				else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "300")
				{
					WinTitle = "[PS_PP361_04] 설계비 내역";
					ReportName = "PS_PP361_04.RPT";
				}
				else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "400")
				{
					WinTitle = "[PS_PP361_06] 외주가공비 내역";
					ReportName = "PS_PP361_06.RPT";
				}
				else if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "500")
				{
					WinTitle = "[PS_PP361_07] 외주제작비 내역";
					ReportName = "PS_PP361_07.RPT";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", oGrid02.DataTable.GetValue(0, i).ToString().Trim()));
				dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", oGrid02.DataTable.GetValue(1, i).ToString().Trim()));
				
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
                   // Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Search")
					{
						if (PS_PP361_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							if (oGrid02.Rows.Count > 0)
							{
								oGrid02.DataTable.Clear();
							}
							if (oGrid03.Rows.Count > 0)
							{
								oGrid03.DataTable.Clear();
							}
							PS_PP361_SearchGrid01Data();
							oForm.Items.Item("Selection").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("SYYYYMM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}

					if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP361_PrintQuery1);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}

					if (pVal.ItemUID == "Btn02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP361_PrintQuery2);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}

					if (pVal.ItemUID == "Btn03")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP361_PrintQuery3);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
					if (pVal.ItemUID == "CardCode")
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처 포맷서치 설정
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
					if (pVal.ItemUID == "Selection" && oGrid02.Rows.Count != 0)
					{
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "100")
						{
							PS_PP361_SearchGrid03Data();
						}
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "200")
						{
							PS_PP361_SearchGrid03Data();
						}
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "300")
						{
							PS_PP361_SearchGrid03Data();
						}
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "400")
						{
							PS_PP361_SearchGrid03Data();
						}
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() == "500")
						{
							PS_PP361_SearchGrid03Data();
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
					if (pVal.ItemUID == "Grid01" && pVal.ColUID == "RowsHeader")
					{
						if (oGrid03.Rows.Count > 0)
						{
							oGrid03.DataTable.Clear();
						}
						if (oForm.Items.Item("Selection").Specific.Selected.Value.ToString().Trim() != "000")
						{
							oForm.Items.Item("Selection").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("SYYYYMM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						PS_PP361_SearchGrid02Data();
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
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					PS_PP361_ResizeForm(FormUID, ref pVal, ref BubbleEvent);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
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
