using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 공정상태 조회
	/// </summary>
	internal class PS_PP551 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.DataTable oDS_PS_PP551A;
		private SAPbouiCOM.DataTable oDS_PS_PP551B;
		private SAPbouiCOM.DataTable oDS_PS_PP551C;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP551.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP551_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP551");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP551_CreateItems();
				PS_PP551_SetComboBox();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Items.Item("Folder01").Specific.Select();  //폼이 로드시 Folder01 선택
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP551_CreateItems
		/// </summary>
		private void PS_PP551_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;

				oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

				oForm.DataSources.DataTables.Add("PS_PP551A");
				oForm.DataSources.DataTables.Add("PS_PP551B");
				oForm.DataSources.DataTables.Add("PS_PP551C");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP551A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP551B");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_PP551C");

				oDS_PS_PP551A = oForm.DataSources.DataTables.Item("PS_PP551A");
				oDS_PS_PP551B = oForm.DataSources.DataTables.Item("PS_PP551B");
				oDS_PS_PP551C = oForm.DataSources.DataTables.Item("PS_PP551C");

				//공정대기 조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode01").Specific.DataBind.SetBound(true, "", "CpCode01");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName01").Specific.DataBind.SetBound(true, "", "CpName01");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode01").Specific.DataBind.SetBound(true, "", "CntcCode01");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName01").Specific.DataBind.SetBound(true, "", "CntcName01");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("WorkGbn01").Specific.DataBind.SetBound(true, "", "WorkGbn01");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType01").Specific.DataBind.SetBound(true, "", "CardType01");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType01").Specific.DataBind.SetBound(true, "", "ItemType01");

				//정렬값
				oForm.DataSources.UserDataSources.Add("OBCol01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OBCol01").Specific.DataBind.SetBound(true, "", "OBCol01");

				//정렬기준
				oForm.DataSources.UserDataSources.Add("OBType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OBType01").Specific.DataBind.SetBound(true, "", "OBType01");

				//공정진행 조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode02").Specific.DataBind.SetBound(true, "", "CpCode02");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName02").Specific.DataBind.SetBound(true, "", "CpName02");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode02").Specific.DataBind.SetBound(true, "", "CntcCode02");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName02").Specific.DataBind.SetBound(true, "", "CntcName02");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("WorkGbn02").Specific.DataBind.SetBound(true, "", "WorkGbn02");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType02").Specific.DataBind.SetBound(true, "", "CardType02");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType02").Specific.DataBind.SetBound(true, "", "ItemType02");

				//정렬값
				oForm.DataSources.UserDataSources.Add("OBCol02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OBCol02").Specific.DataBind.SetBound(true, "", "OBCol02");

				//정렬기준
				oForm.DataSources.UserDataSources.Add("OBType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OBType02").Specific.DataBind.SetBound(true, "", "OBType02");

				//공정완료 조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID03").Specific.DataBind.SetBound(true, "", "BPLID03");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt03").Specific.DataBind.SetBound(true, "", "FrDt03");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt03").Specific.DataBind.SetBound(true, "", "ToDt03");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode03").Specific.DataBind.SetBound(true, "", "CpCode03");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName03").Specific.DataBind.SetBound(true, "", "CpName03");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode03").Specific.DataBind.SetBound(true, "", "CntcCode03");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName03").Specific.DataBind.SetBound(true, "", "CntcName03");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("WorkGbn03").Specific.DataBind.SetBound(true, "", "WorkGbn03");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType03").Specific.DataBind.SetBound(true, "", "CardType03");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType03").Specific.DataBind.SetBound(true, "", "ItemType03");

				oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt03").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt03").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP551_SetComboBox
		/// </summary>
		private void PS_PP551_SetComboBox()
		{
			string sQry;
			string User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				User_BPLId = dataHelpClass.User_BPLID();

				//공정대기 조회
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);

				//작업구분
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("WorkGbn01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn01").Specific, sQry, "%", false, false);

				//거래처구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				oForm.Items.Item("CardType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType01").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				oForm.Items.Item("ItemType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType01").Specific, sQry, "%", false, false);

				//정렬값
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P209'";
				oForm.Items.Item("OBCol01").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OBCol01").Specific, sQry, "%", false, false);

				//정렬기준
				oForm.Items.Item("OBType01").Specific.ValidValues.Add("0", "오름차순");
				oForm.Items.Item("OBType01").Specific.ValidValues.Add("1", "내림차순");
				oForm.Items.Item("OBType01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//공정진행 조회
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);

				//작업구분
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("WorkGbn02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn02").Specific, sQry, "%", false, false);

				//거래처구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				oForm.Items.Item("CardType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType02").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				oForm.Items.Item("ItemType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType02").Specific, sQry, "%", false, false);

				//정렬값
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P209'";
				oForm.Items.Item("OBCol02").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OBCol02").Specific, sQry, "%", false, false);

				//정렬기준
				oForm.Items.Item("OBType02").Specific.ValidValues.Add("0", "오름차순");
				oForm.Items.Item("OBType02").Specific.ValidValues.Add("1", "내림차순");
				oForm.Items.Item("OBType02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//공정완료 조회
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID03").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);

				//작업구분
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("WorkGbn03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn03").Specific, sQry, "%", false, false);

				//거래처구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				oForm.Items.Item("CardType03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType03").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				oForm.Items.Item("ItemType03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType03").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP551_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP551_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oUID == "CntcCode01")
				{
					oForm.Items.Item("CntcName01").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CntcCode02")
				{
					oForm.Items.Item("CntcName02").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CntcCode03")
				{
					oForm.Items.Item("CntcName03").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode01")
				{
					oForm.Items.Item("CpName01").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode02")
				{
					oForm.Items.Item("CpName02").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
				else if (oUID == "CpCode03")
				{
					oForm.Items.Item("CpName03").Specific.Value = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP551_ResizeForm
		/// </summary>
		private void PS_PP551_ResizeForm()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 120;
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
		/// PS_PP551_MTX01  공정대기조회
		/// </summary>
		private void PS_PP551_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLID;    //사업장
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string CpCode;   //공정
			string CntcCode; //등록자사번
			string WorkGbn;  //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode01").Specific.Value.ToString().Trim(); ;
				CntcCode = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				WorkGbn = oForm.Items.Item("WorkGbn01").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType01").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType01").Specific.Selected.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC PS_PP551_01 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += CntcCode + "','";
				sQry += WorkGbn + "','";
				sQry += CardType + "','";
				sQry += ItemType + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP551A.ExecuteQuery(sQry);

				oGrid01.Columns.Item(4).RightJustified = true;
				oGrid01.Columns.Item(10).RightJustified = true;

				if (oGrid01.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
				oGrid01.AutoResizeColumns();
				oForm.Update();
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP551_MTX02  공정시작조회
		/// </summary>
		private void PS_PP551_MTX02()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLID;    //사업장
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string CpCode;   //공정
			string CntcCode; //등록자사번
			string WorkGbn;  //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID    = oForm.Items.Item("BPLID02").Specific.Selected.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn02").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Selected.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC PS_PP551_02 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += CntcCode + "','";
				sQry += WorkGbn + "','";
				sQry += CardType + "','";
				sQry += ItemType + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP551B.ExecuteQuery(sQry);

				oGrid02.Columns.Item(4).RightJustified = true;
				oGrid02.Columns.Item(10).RightJustified = true;

				if (oGrid02.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
				oGrid02.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP551_MTX03  공정완료조회
		/// </summary>
		private void PS_PP551_MTX03()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLID;    //사업장
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string CpCode;   //공정
			string CntcCode; //등록자사번
			string WorkGbn;  //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID    = oForm.Items.Item("BPLID03").Specific.Selected.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt03").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt03").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode03").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn03").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType03").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType03").Specific.Selected.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC PS_PP551_03 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CpCode + "','";
				sQry += CntcCode + "','";
				sQry += WorkGbn + "','";
				sQry += CardType + "','";
				sQry += ItemType + "'";
				
				oGrid03.DataTable.Clear();
				oDS_PS_PP551C.ExecuteQuery(sQry);

				oGrid03.Columns.Item(4).RightJustified = true;
				oGrid03.Columns.Item(10).RightJustified = true;

				if (oGrid03.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
				oGrid03.AutoResizeColumns();
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP551_PrintReport01  공정대기리포트 출력
		/// </summary>
		[STAThread]
		private void PS_PP551_PrintReport01()
		{
			string WinTitle;
			string ReportName;
			string BPLID;	 //사업장
			string FrDt;	 //기간(Fr)
			string ToDt;	 //기간(To)
			string CpCode;	 //공정
			string CntcCode; //등록자사번
			string WorkGbn;	 //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분
			string OBCol;	 //정렬값
			string OBType;	 //정렬기준

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID    = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode01").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn01").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType01").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType01").Specific.Selected.Value.ToString().Trim();
				OBCol    = oForm.Items.Item("OBCol01").Specific.Selected.Value.ToString().Trim();
				OBType   = oForm.Items.Item("OBType01").Specific.Selected.Value.ToString().Trim();

				WinTitle = "[PS_PP551] 레포트";
				ReportName = "PS_PP551_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkGbn", WorkGbn));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkerCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrderByColumn", OBCol));
				dataPackParameter.Add(new PSH_DataPackClass("@OrderByType", OBType));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP551_PrintReport02  공정시작리포트 출력
		/// </summary>
		[STAThread]
		private void PS_PP551_PrintReport02()
		{

			string WinTitle;
			string ReportName;
			string BPLID;    //사업장
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string CpCode;   //공정
			string CntcCode; //등록자사번
			string WorkGbn;  //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분
			string OBCol;    //정렬값
			string OBType;   //정렬기준
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID    = oForm.Items.Item("BPLID02").Specific.Selected.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn02").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Selected.Value.ToString().Trim();
				OBCol    = oForm.Items.Item("OBCol02").Specific.Selected.Value.ToString().Trim();
				OBType   = oForm.Items.Item("OBType02").Specific.Selected.Value.ToString().Trim();

				WinTitle = "[PS_PP551] 레포트";
				ReportName = "PS_PP551_02.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkGbn", WorkGbn));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkerCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrderByColumn", OBCol));
				dataPackParameter.Add(new PSH_DataPackClass("@OrderByType", OBType));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP551_PrintReport03  공정완료리포트 출력
		/// </summary>
		[STAThread]
		private void PS_PP551_PrintReport03()
		{
			string WinTitle;
			string ReportName;
			string BPLID;    //사업장
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string CpCode;   //공정
			string CntcCode; //등록자사번
			string WorkGbn;  //작업구분
			string CardType; //거래처구분
			string ItemType; //품목구분
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID    = oForm.Items.Item("BPLID03").Specific.Selected.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt03").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt03").Specific.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode03").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn03").Specific.Selected.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType03").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType03").Specific.Selected.Value.ToString().Trim();

				WinTitle = "[PS_PP551] 레포트";
				ReportName = "PS_PP551_03.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkGbn", WorkGbn));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkerCode", CntcCode));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP551_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP551_MTX02();
						}
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP551_MTX03();
						}
					}
					else if (pVal.ItemUID == "BtnPrt01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP551_PrintReport01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP551_PrintReport02);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP551_PrintReport03);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					else if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					else if (pVal.ItemUID == "Folder03")
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode01", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode02", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode03", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode03", "");
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
					PS_PP551_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Grid01")
					{
						if (pVal.Row == -1)
						{
							oGrid01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid01.Rows.SelectedRows.Count > 0)
							{
							}
							else
							{
								BubbleEvent = false;
							}
						}
					}
					else if (pVal.ItemUID == "Grid02")
					{
						if (pVal.Row == -1)
						{
							oGrid02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid02.Rows.SelectedRows.Count > 0)
							{
							}
							else
							{
								BubbleEvent = false;
							}
						}
					}
					else if (pVal.ItemUID == "Grid03")
					{
						if (pVal.Row == -1)
						{
							oGrid03.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid03.Rows.SelectedRows.Count > 0)
							{
							}
							else
							{
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
						PS_PP551_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_PP551_ResizeForm();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP551A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP551B);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP551C);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
