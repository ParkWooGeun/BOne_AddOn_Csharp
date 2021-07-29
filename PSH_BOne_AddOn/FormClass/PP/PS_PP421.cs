using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 비가동내역조회
	/// </summary>
	internal class PS_PP421 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.DataTable oDS_PS_PP421L;
		private SAPbouiCOM.DataTable oDS_PS_PP421M;
		private SAPbouiCOM.DataTable oDS_PS_PP421N;

		/// <summary>
		/// 화면 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP421.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP421_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP421");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP421_CreateItems();
				PS_PP421_SetComboBox();
				PS_PP421_EnableFormItem();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP421_CreateItems
		/// </summary>
		private void PS_PP421_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;

				oForm.DataSources.DataTables.Add("PS_PP421L");
				oForm.DataSources.DataTables.Add("PS_PP421M");
				oForm.DataSources.DataTables.Add("PS_PP421N");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP421L");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP421M");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_PP421N");

				oDS_PS_PP421L = oForm.DataSources.DataTables.Item("PS_PP421L");
				oDS_PS_PP421M = oForm.DataSources.DataTables.Item("PS_PP421M");
				oDS_PS_PP421N = oForm.DataSources.DataTables.Item("PS_PP421N");

				//인원별 비가동현황
				//사업장1
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

				//인원별 비가동별 세부현황
				//사업장2
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//팀2
				oForm.DataSources.UserDataSources.Add("TeamCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TeamCode02").Specific.DataBind.SetBound(true, "", "TeamCode02");

				//담당2
				oForm.DataSources.UserDataSources.Add("RspCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("RspCode02").Specific.DataBind.SetBound(true, "", "RspCode02");

				//반2
				oForm.DataSources.UserDataSources.Add("ClsCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ClsCode02").Specific.DataBind.SetBound(true, "", "ClsCode02");

				//기간(시작)2
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//기간(종료)2
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");

				//작업자2
				oForm.DataSources.UserDataSources.Add("WorkCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("WorkCode02").Specific.DataBind.SetBound(true, "", "WorkCode02");

				//작업자성명2
				oForm.DataSources.UserDataSources.Add("WorkName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkName02").Specific.DataBind.SetBound(true, "", "WorkName02");

				//비가동코드2
				oForm.DataSources.UserDataSources.Add("CsCpCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("CsCpCode02").Specific.DataBind.SetBound(true, "", "CsCpCode02");

				//비가동명2
				oForm.DataSources.UserDataSources.Add("CsCpName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("CsCpName02").Specific.DataBind.SetBound(true, "", "CsCpName02");

				//월별 비가동현황
				//사업장3
				oForm.DataSources.UserDataSources.Add("BPLID03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID03").Specific.DataBind.SetBound(true, "", "BPLID03");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM03").Specific.DataBind.SetBound(true, "", "StdYM03");

				//작업자3
				oForm.DataSources.UserDataSources.Add("WorkCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("WorkCode03").Specific.DataBind.SetBound(true, "", "WorkCode03");

				//팀3
				oForm.DataSources.UserDataSources.Add("TeamCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TeamCode03").Specific.DataBind.SetBound(true, "", "TeamCode03");

				//작업자성명3
				oForm.DataSources.UserDataSources.Add("WorkName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WorkName03").Specific.DataBind.SetBound(true, "", "WorkName03");

				//날짜기본SET
				oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("StdYM03").Specific.Value = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP421_SetComboBox
		/// </summary>
		private void PS_PP421_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false); //사업장1
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false); //사업장2
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID03").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false); //사업장3
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP421_EnableFormItem
		/// </summary>
		private void PS_PP421_EnableFormItem()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLID02").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
					PS_PP421_FlushToItemValue("BPLID02", 0, ""); //팀, 담당, 반 콤보박스 강제 설정
					PS_PP421_FlushToItemValue("BPLID03", 0, ""); //팀, 담당, 반 콤보박스 강제 설정
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
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
		/// PS_PP421_ResizeForm
		/// </summary>
		private void PS_PP421_ResizeForm()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 75;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				if (oGrid01.Columns.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}

				if (oGrid02.Columns.Count > 0)
				{
					oGrid02.AutoResizeColumns();
				}

				if (oGrid03.Columns.Count > 0)
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
		/// PS_PP421_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP421_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;
			string BPLId;
			string TeamCode;
			string RspCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "BPLID02":
						BPLId = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
						if (oForm.Items.Item("TeamCode02").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode02").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode02").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//부서콤보세팅
						oForm.Items.Item("TeamCode02").Specific.ValidValues.Add("%", "전체");
						sQry = "   SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '1'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char2 = '" + BPLId + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

						oForm.Items.Item("RspCode02").Specific.ValidValues.Add("%", "전체");
						oForm.Items.Item("RspCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						oForm.Items.Item("ClsCode02").Specific.ValidValues.Add("%", "전체");
						oForm.Items.Item("ClsCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "TeamCode02":
						TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("RspCode02").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("RspCode02").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("RspCode02").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//담당콤보세팅
						oForm.Items.Item("RspCode02").Specific.ValidValues.Add("%", "전체");
						sQry = "   SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '2'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "RspCode02":
						TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
						RspCode = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("ClsCode02").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("ClsCode02").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("ClsCode02").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//반콤보세팅
						oForm.Items.Item("ClsCode02").Specific.ValidValues.Add("%", "전체");
						sQry = "   SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '9'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + RspCode + "'";
						sQry += "                 AND U_Char2 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("ClsCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "WorkCode02":
						oForm.Items.Item("WorkName02").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("WorkCode02").Specific.Value.ToString().Trim() + "'", ""); //성명
						break;

					case "CsCpCode02":
						oForm.Items.Item("CsCpName02").Specific.Value = dataHelpClass.Get_ReData("U_CdName", "U_Minor", "[@PS_SY001L]", "'" + oForm.Items.Item("CsCpCode02").Specific.Value.ToString().Trim() + "'", " AND Code = 'P005'");	//비가동명
						break;

					case "WorkCode03":
						oForm.Items.Item("WorkName03").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("WorkCode03").Specific.Value.ToString().Trim() + "'", ""); //성명
						break;

					case "BPLID03":
						BPLId = oForm.Items.Item("BPLID03").Specific.Value.ToString().Trim();
						if (oForm.Items.Item("TeamCode03").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode03").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode03").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//부서콤보세팅
						oForm.Items.Item("TeamCode03").Specific.ValidValues.Add("%", "전체");
						sQry = "   SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '1'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char2 = '" + BPLId + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode03").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP421_MTX01
		/// </summary>
		private void PS_PP421_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLId;
			string FrDt;
			string ToDt;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				FrDt  = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt  = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				sQry = "EXEC PS_PP421_01 '";
				sQry += BPLId + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP421L.ExecuteQuery(sQry);

				oGrid01.Columns.Item(11).RightJustified = true;

				if (oGrid01.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid01.AutoResizeColumns();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP421_MTX02
		/// </summary>
		private void PS_PP421_MTX02()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLId;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;	 //반
			string FrDt;	 //기간(시작)
			string ToDt;	 //기간(종료)
			string WorkCode; //작업자
			string CsCpCode; //비가동코드
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
		
				BPLId = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode02").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode02").Specific.Value.ToString().Trim();
				CsCpCode = oForm.Items.Item("CsCpCode02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				sQry = "EXEC PS_PP421_02 '";
				sQry += BPLId + "','";
				sQry += TeamCode + "','";
				sQry += RspCode + "','";
				sQry += ClsCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += WorkCode + "','";
				sQry += CsCpCode + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP421M.ExecuteQuery(sQry);

				oGrid02.Columns.Item(5).RightJustified = true;
				oGrid02.Columns.Item(6).RightJustified = true;
				oGrid02.Columns.Item(7).RightJustified = true;
				oGrid02.Columns.Item(8).RightJustified = true;
				oGrid02.Columns.Item(9).RightJustified = true;
				oGrid02.Columns.Item(10).RightJustified = true;
				oGrid02.Columns.Item(11).RightJustified = true;
				oGrid02.Columns.Item(12).RightJustified = true;
				oGrid02.Columns.Item(13).RightJustified = true;
				oGrid02.Columns.Item(14).RightJustified = true;
				oGrid02.Columns.Item(15).RightJustified = true;
				oGrid02.Columns.Item(16).RightJustified = true;
				oGrid02.Columns.Item(17).RightJustified = true;
				oGrid02.Columns.Item(18).RightJustified = true;
				oGrid02.Columns.Item(19).RightJustified = true;
				oGrid02.Columns.Item(20).RightJustified = true;
				oGrid02.Columns.Item(21).RightJustified = true;
				oGrid02.Columns.Item(22).RightJustified = true;
				oGrid02.Columns.Item(23).RightJustified = true;
				oGrid02.Columns.Item(24).RightJustified = true;
				oGrid02.Columns.Item(25).RightJustified = true;
				oGrid02.Columns.Item(26).RightJustified = true;
				oGrid02.Columns.Item(27).RightJustified = true;
				oGrid02.Columns.Item(28).RightJustified = true;
				oGrid02.Columns.Item(29).RightJustified = true;
				oGrid02.Columns.Item(30).RightJustified = true;
				oGrid02.Columns.Item(31).RightJustified = true;
				oGrid02.Columns.Item(32).RightJustified = true;
				oGrid02.Columns.Item(33).RightJustified = true;
				oGrid02.Columns.Item(34).RightJustified = true;
				oGrid02.Columns.Item(35).RightJustified = true;

				if (oGrid02.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid02.AutoResizeColumns();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP421_MTX03
		/// </summary>
		private void PS_PP421_MTX03()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLId;
			string StdYM; //기준년월
			string WorkCode;
			string TeamCode;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLID03").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM03").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode03").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode03").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				sQry = "EXEC PS_PP421_03 '";
				sQry += BPLId + "','";
				sQry += StdYM + "','";
				sQry += TeamCode + "','";
				sQry += WorkCode + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_PP421N.ExecuteQuery(sQry);

				oGrid03.Columns.Item(3).RightJustified = true;
				oGrid03.Columns.Item(4).RightJustified = true;
				oGrid03.Columns.Item(5).RightJustified = true;
				oGrid03.Columns.Item(6).RightJustified = true;
				oGrid03.Columns.Item(7).RightJustified = true;
				oGrid03.Columns.Item(8).RightJustified = true;
				oGrid03.Columns.Item(9).RightJustified = true;
				oGrid03.Columns.Item(10).RightJustified = true;
				oGrid03.Columns.Item(11).RightJustified = true;
				oGrid03.Columns.Item(12).RightJustified = true;
				oGrid03.Columns.Item(13).RightJustified = true;
				oGrid03.Columns.Item(14).RightJustified = true;
				oGrid03.Columns.Item(15).RightJustified = true;
				oGrid03.Columns.Item(16).RightJustified = true;
				oGrid03.Columns.Item(17).RightJustified = true;
				oGrid03.Columns.Item(18).RightJustified = true;
				oGrid03.Columns.Item(19).RightJustified = true;
				oGrid03.Columns.Item(20).RightJustified = true;
				oGrid03.Columns.Item(21).RightJustified = true;
				oGrid03.Columns.Item(22).RightJustified = true;
				oGrid03.Columns.Item(23).RightJustified = true;
				oGrid03.Columns.Item(24).RightJustified = true;
				oGrid03.Columns.Item(25).RightJustified = true;
				oGrid03.Columns.Item(26).RightJustified = true;
				oGrid03.Columns.Item(27).RightJustified = true;
				oGrid03.Columns.Item(28).RightJustified = true;
				oGrid03.Columns.Item(29).RightJustified = true;
				oGrid03.Columns.Item(30).RightJustified = true;
				oGrid03.Columns.Item(31).RightJustified = true;
				oGrid03.Columns.Item(32).RightJustified = true;
				oGrid03.Columns.Item(33).RightJustified = true;
				oGrid03.Columns.Item(34).RightJustified = true;

				if (oGrid03.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid03.AutoResizeColumns();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP421_PrintReport02
		/// </summary>
		[STAThread]
		private void PS_PP421_PrintReport02()
		{
			string WinTitle;
			string ReportName;
			string BPLId;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;	 //반
			string FrDt;	 //기간(시작)
			string ToDt;	 //기간(종료)
			string WorkCode; //작업자
			string CsCpCode; //비가동코드
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode02").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode02").Specific.Value.ToString().Trim();
				CsCpCode = oForm.Items.Item("CsCpCode02").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP421] 레포트";
				ReportName = "PS_PP421_02.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkCode", WorkCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CsCpCode", CsCpCode));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP421_PrintReport03
		/// </summary>
		[STAThread]
		private void PS_PP421_PrintReport03()
		{
			string WinTitle;
			string ReportName;
			string BPLId;
			string StdYM; //기준년월
			string WorkCode;
			string TeamCode;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLID03").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM03").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode03").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode03").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP421] 레포트";
				ReportName = "PS_PP421_03.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkCode", WorkCode));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));

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
                   // Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSrch01")
					{
						PS_PP421_MTX01();
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						PS_PP421_MTX02();
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						PS_PP421_MTX03();
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP421_PrintReport02);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt03")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP421_PrintReport03);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WorkCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CsCpCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WorkCode03", "");
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
					PS_PP421_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_PP421_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
				}
				oForm.Freeze(false);
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
					PS_PP421_ResizeForm();
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
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP421L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP421M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP421N);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
