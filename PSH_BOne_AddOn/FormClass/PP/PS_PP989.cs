using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기간별 표준공수 대비 실동공수 조회
	/// </summary>
	internal class PS_PP989 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DBDataSource oDS_PS_PP989O;
		private SAPbouiCOM.DataTable oDS_PS_PP989L;

		/// <summary>
		/// 화면 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP989.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP989_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP989");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP989_CreateItems();
				PS_PP989_ComboBox_Setting();
				PS_PP989_FormItemEnabled();
				PS_PP989_FormResize();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);

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
		/// PS_PP989_CreateItems
		/// </summary>
		private void PS_PP989_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_PP989L");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP989L");
				oDS_PS_PP989L = oForm.DataSources.DataTables.Item("PS_PP989L");
				oDS_PS_PP989O = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				oMat = oForm.Items.Item("Mat02").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat.AutoResizeColumns();

				//작번별 집계
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기간(Fr)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

				//기간(To)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

				//작업구분
				oForm.DataSources.UserDataSources.Add("OrdGbn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdGbn01").Specific.DataBind.SetBound(true, "", "OrdGbn01");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("OrdNum01").Specific.DataBind.SetBound(true, "", "OrdNum01");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("OrdSub101", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("OrdSub101").Specific.DataBind.SetBound(true, "", "OrdSub101");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("OrdSub201", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("OrdSub201").Specific.DataBind.SetBound(true, "", "OrdSub201");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName01").Specific.DataBind.SetBound(true, "", "ItemName01");

				//담당자사번
				oForm.DataSources.UserDataSources.Add("CntcCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode01").Specific.DataBind.SetBound(true, "", "CntcCode01");

				//담당자성명
				oForm.DataSources.UserDataSources.Add("CntcName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcName01").Specific.DataBind.SetBound(true, "", "CntcName01");

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardCode01").Specific.DataBind.SetBound(true, "", "CardCode01");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CardName01").Specific.DataBind.SetBound(true, "", "CardName01");

				//공정별 집계
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//팀
				oForm.DataSources.UserDataSources.Add("TeamCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("TeamCode02").Specific.DataBind.SetBound(true, "", "TeamCode02");

				//담당
				oForm.DataSources.UserDataSources.Add("RspCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("RspCode02").Specific.DataBind.SetBound(true, "", "RspCode02");

				//소속반
				oForm.DataSources.UserDataSources.Add("ClsCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ClsCode02").Specific.DataBind.SetBound(true, "", "ClsCode02");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType02").Specific.DataBind.SetBound(true, "", "CardType02");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType02").Specific.DataBind.SetBound(true, "", "ItemType02");

				//생산완료여부
				oForm.DataSources.UserDataSources.Add("WCYN02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("WCYN02").Specific.DataBind.SetBound(true, "", "WCYN02");

				//일자기준
				oForm.DataSources.UserDataSources.Add("DateStd02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("DateStd02").Specific.DataBind.SetBound(true, "", "DateStd02");

				//기간(Fr)
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//기간(To)
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");

				//품목(작번)
				oForm.DataSources.UserDataSources.Add("ItemCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode02").Specific.DataBind.SetBound(true, "", "ItemCode02");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName02").Specific.DataBind.SetBound(true, "", "ItemName02");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemSpec02").Specific.DataBind.SetBound(true, "", "ItemSpec02");

				//공정
				oForm.DataSources.UserDataSources.Add("CpCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode02").Specific.DataBind.SetBound(true, "", "CpCode02");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CpName02").Specific.DataBind.SetBound(true, "", "CpName02");

				//표준공수계(시간)
				oForm.DataSources.UserDataSources.Add("TStdTime02", SAPbouiCOM.BoDataType.dt_QUANTITY);
				oForm.Items.Item("TStdTime02").Specific.DataBind.SetBound(true, "", "TStdTime02");

				//표준공수계(금액)
				oForm.DataSources.UserDataSources.Add("TStdAmt02", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("TStdAmt02").Specific.DataBind.SetBound(true, "", "TStdAmt02");

				//실동공수계(시간)
				oForm.DataSources.UserDataSources.Add("TWkTime02", SAPbouiCOM.BoDataType.dt_QUANTITY);
				oForm.Items.Item("TWkTime02").Specific.DataBind.SetBound(true, "", "TWkTime02");

				//실동공수계(금액)
				oForm.DataSources.UserDataSources.Add("TWkAmt02", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("TWkAmt02").Specific.DataBind.SetBound(true, "", "TWkAmt02");

				//작업구분
				oForm.DataSources.UserDataSources.Add("OrdGbn02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdGbn02").Specific.DataBind.SetBound(true, "", "OrdGbn02");


				//날짜 기본SET
				oForm.Items.Item("FrDt01").Specific.VALUE = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt01").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.VALUE = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_ComboBox_Setting
		/// </summary>
		private void PS_PP989_ComboBox_Setting()
		{
			string sQry;
			string BPLID;
			
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//작번별 집계
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//작업구분
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("OrdGbn01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn01").Specific, sQry, "", false, false);
				oForm.Items.Item("OrdGbn01").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);


				//공정별 집계
				//사업장2
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

				//거래처구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				oForm.Items.Item("CardType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType02").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType02").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품목구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				oForm.Items.Item("ItemType02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType02").Specific, sQry, "", false, false);
				oForm.Items.Item("ItemType02").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//생산완료여부
				oForm.Items.Item("WCYN02").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("WCYN02").Specific.ValidValues.Add("B", "미완료");
				oForm.Items.Item("WCYN02").Specific.ValidValues.Add("C", "완료");
				oForm.Items.Item("WCYN02").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//일자기준
				oForm.Items.Item("DateStd02").Specific.ValidValues.Add("01", "작업지시");
				oForm.Items.Item("DateStd02").Specific.ValidValues.Add("02", "작업일보");
				oForm.Items.Item("DateStd02").Specific.Select("02", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//작업구분
				sQry = " SELECT      Code, ";
				sQry += "             Name ";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				oForm.Items.Item("OrdGbn02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn02").Specific, sQry, "", false, false);
				oForm.Items.Item("OrdGbn02").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_FormItemEnabled
		/// </summary>
		private void PS_PP989_FormItemEnabled()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLID02").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
					PS_PP989_FlushToItemValue("BPLID02", 0, ""); //팀, 담당, 반 콤보박스 강제 설정
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP989_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;

			string BPLID;
			string TeamCode;
			string RspCode;

			string OrdNum;
			string SubNo1;
			string SubNo2;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "BPLID02":
						BPLID = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("TeamCode02").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode02").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode02").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//부서콤보세팅
						oForm.Items.Item("TeamCode02").Specific.ValidValues.Add("%", "전체");
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "             U_CodeNm As [Name]";
						sQry += " FROM        [@PS_HR200L]";
						sQry += " WHERE       Code = '1'";
						sQry += "             AND U_UseYN = 'Y'";
						sQry += "             AND U_Char2 = '" + BPLID + "'";
						sQry += " ORDER BY    U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "             U_CodeNm As [Name]";
						sQry += " FROM        [@PS_HR200L]";
						sQry += " WHERE       Code = '2'";
						sQry += "             AND U_UseYN = 'Y'";
						sQry += "             AND U_Char1 = '" + TeamCode + "'";
						sQry += " ORDER BY    U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "RspCode02":
						TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
						RspCode  = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("ClsCode02").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("ClsCode02").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("ClsCode02").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//반콤보세팅
						oForm.Items.Item("ClsCode02").Specific.ValidValues.Add("%", "전체");
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "             U_CodeNm As [Name]";
						sQry += " FROM        [@PS_HR200L]";
						sQry += " WHERE       Code = '9'";
						sQry += "             AND U_UseYN = 'Y'";
						sQry += "             AND U_Char1 = '" + RspCode + "'";
						sQry += "             AND U_Char2 = '" + TeamCode + "'";
						sQry += " ORDER BY    U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode02").Specific, sQry, "", false, false);
						oForm.Items.Item("ClsCode02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "ItemCode02":
						sQry = " SELECT      FrgnName,";
						sQry += "             U_Size";
						sQry += " FROM        OITM";
						sQry += " WHERE       ItemCode = '" + oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("ItemName02").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						oForm.Items.Item("ItemSpec02").Specific.VALUE = oRecordSet.Fields.Item(1).Value.ToString().Trim();
						break;

					case "CpCode02":
						sQry = "        SELECT      U_CpName";
						sQry += " FROM        [@PS_PP001L]";
						sQry += " WHERE       U_CpCode = '" + oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("CpName02").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;

					case "CntcCode01":
						sQry = " SELECT      U_FullName";
						sQry += " FROM        [@PH_PY001A]";
						sQry += " WHERE       Code = '" + oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("CntcName01").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;

					case "CardCode01":
						sQry = "        SELECT      CardName";
						sQry += " FROM        [OCRD]";
						sQry += " WHERE       CardCode = '" + oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("CardName01").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}

				if (oUID == "OrdNum01" || oUID == "OrdSub101" || oUID == "OrdSub201")
				{
					OrdNum = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();
					SubNo1 = oForm.Items.Item("OrdSub101").Specific.Value.ToString().Trim();
					SubNo2 = oForm.Items.Item("OrdSub201").Specific.Value.ToString().Trim();

					sQry = " SELECT      CASE";
					sQry += "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "                 ELSE T0.U_JakMyung";
					sQry += "             END AS [ItemName],";
					sQry += "             CASE";
					sQry += "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "                 ELSE T0.U_JakSize";
					sQry += "             END AS [SPEC]";
					sQry += " FROM        [@PS_PP020H] AS T0";
					sQry += " WHERE       T0.U_JakName = '" + OrdNum + "'";
					sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + SubNo1 + "' = '' THEN '00' ELSE '" + SubNo1 + "' END";
					sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + SubNo2 + "' = '' THEN '000' ELSE '" + SubNo2 + "' END";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName01").Specific.VALUE = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
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
		/// PS_PP989_Add_MatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP989_Add_MatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP989O.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP989O.Offset = oRow;
				oDS_PS_PP989O.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_FormResize
		/// </summary>
		private void PS_PP989_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 85;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				if (oGrid.Columns.Count > 0)
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
		/// PS_PP989_GetDetail
		/// </summary>
		/// <param name="pCpCode"></param>
		/// <param name="pCpName"></param>
		private void PS_PP989_GetDetail(string pCpCode, string pCpName)
		{
			string BPLID;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;   //반
			string CardType; //거래처구분
			string ItemType; //품목구분
			string WCYN;	 //생산완료여부
			string DateStd;	 //일자기준
			string FrDt;	 //기간(Fr)
			string ToDt;	 //기간(To)
			string ItemCode; //품목코드(작번)
			string CpCode;	 //공정
			string CpName;	 //공정명
			string OrdGbn;   //작업구분

			try
			{
				BPLID    = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
				RspCode  = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();
				ClsCode  = oForm.Items.Item("ClsCode02").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Value.ToString().Trim();
				WCYN     = oForm.Items.Item("WCYN02").Specific.Value.ToString().Trim();
				DateStd  = oForm.Items.Item("DateStd02").Specific.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim();
				CpCode   = pCpCode;
				CpName   = pCpName;
				OrdGbn   = oForm.Items.Item("OrdGbn02").Specific.Value.ToString().Trim();

				PS_PP990 oTempClass = new PS_PP990();
				oTempClass.LoadForm(BPLID, TeamCode, RspCode, ClsCode, CardType, ItemType, WCYN, DateStd, FrDt, ToDt,ItemCode, CpCode, CpName, OrdGbn);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_MTX01
		/// </summary>
		private void PS_PP989_MTX01()
		{
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string FrDt;
			string ToDt;
			string OrdGbn;
			string OrdNum;   //작번
			string OrdSub1;	 //서브작번1
			string OrdSub2;	 //서브작번2
			string CntcCode; //담당
			string CardCode; //거래처

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID    = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				OrdGbn   = oForm.Items.Item("OrdGbn01").Specific.Value.ToString().Trim();
				OrdNum   = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();
				OrdSub1  = oForm.Items.Item("OrdSub101").Specific.Value.ToString().Trim();
				OrdSub2  = oForm.Items.Item("OrdSub201").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				sQry = " EXEC PS_PP989_01 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += OrdGbn + "','";
				sQry += OrdNum + "','";
				sQry += OrdSub1 + "','";
				sQry += OrdSub2 + "','";
				sQry += CntcCode + "','";
				sQry += CardCode + "'";

				oGrid.DataTable.Clear();
				oDS_PS_PP989L.ExecuteQuery(sQry);

				oGrid.Columns.Item(16).RightJustified = true;
				oGrid.Columns.Item(17).RightJustified = true;

				if (oGrid.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid.AutoResizeColumns();
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP989_MTX02
		/// </summary>
		private void PS_PP989_MTX02()
		{
			int i;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;	 //반
			string CardType; //거래처구분
			string ItemType; //품목구분
			string WCYN;	 //생산완료여부
			string DateStd;	 //일자기준
			string FrDt;	 //기간(Fr)
			string ToDt;	 //기간(To)
			string ItemCode; //품목코드(작번)
			string CpCode;	 //공정
			string OrdGbn;	 //작업구분
			string CntcCode; //조회자 사번
							 
			double TStdTime = 0; //표준공수계(시간)
			decimal TStdAmt = 0; //표준공수계(금액)
			double TWkTime = 0;	 //실동공수계(시간)
			decimal TWkAmt = 0;  //실동공수계(금액)

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode02").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Value.ToString().Trim();
				WCYN = oForm.Items.Item("WCYN02").Specific.Value.ToString().Trim();
				DateStd = oForm.Items.Item("DateStd02").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim();
				OrdGbn = oForm.Items.Item("OrdGbn02").Specific.Value.ToString().Trim();
				CntcCode = dataHelpClass.User_MSTCOD();

				ProgressBar01.Text = "조회시작!";

				sQry = "      EXEC [PS_PP989_02] '";
				sQry = sQry + BPLID + "','";
				sQry = sQry + TeamCode + "','";
				sQry = sQry + RspCode + "','";
				sQry = sQry + ClsCode + "','";
				sQry = sQry + CardType + "','";
				sQry = sQry + ItemType + "','";
				sQry = sQry + WCYN + "','";
				sQry = sQry + DateStd + "','";
				sQry = sQry + FrDt + "','";
				sQry = sQry + ToDt + "','";
				sQry = sQry + ItemCode + "','";
				sQry = sQry + CpCode + "','";
				sQry = sQry + OrdGbn + "','";
				sQry = sQry + CntcCode + "'";

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP989O.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP989_Add_MatrixRow01(0, true);
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{

					if (i + 1 > oDS_PS_PP989O.Size)
					{
						oDS_PS_PP989O.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP989O.Offset = i;

					oDS_PS_PP989O.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP989O.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	//공정코드
					oDS_PS_PP989O.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	//공정명
					oDS_PS_PP989O.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());	//작업일자
					oDS_PS_PP989O.SetValue("U_ColPrc01", i, oRecordSet.Fields.Item("Price").Value.ToString().Trim());	//공정단가
					oDS_PS_PP989O.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("StdTime").Value.ToString().Trim());	//표준공수
					oDS_PS_PP989O.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("StdAmt").Value.ToString().Trim());	//표준공수금액
					oDS_PS_PP989O.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("WkTime").Value.ToString().Trim());	//실동공수
					oDS_PS_PP989O.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("WkAmt").Value.ToString().Trim());   //실동공수금액

					if (oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[소 계]" || oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[합 계]")
					{
						TStdTime += 0;
					}
					else
					{
						TStdTime += Convert.ToDouble(oRecordSet.Fields.Item("StdTime").Value.ToString().Trim());
					}
					if (oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[소 계]" || oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[합 계]")
					{
						TStdAmt += 0;
					}
					else
					{
						TStdAmt += Convert.ToDecimal(oRecordSet.Fields.Item("StdAmt").Value.ToString().Trim());
					}
					if (oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[소 계]" || oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[합 계]")
					{
						TWkTime += 0;
					}
					else
					{
						TWkTime += Convert.ToDouble(oRecordSet.Fields.Item("WkTime").Value.ToString().Trim());
					}
					if (oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[소 계]" || oRecordSet.Fields.Item("DocDate").Value.ToString().Trim() == "[합 계]")
					{
						TWkAmt += 0;
					}
					else
					{
						TWkAmt += Convert.ToDecimal(oRecordSet.Fields.Item("WkAmt").Value.ToString().Trim());
					}

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();

				oForm.Items.Item("TStdTime02").Specific.VALUE = TStdTime; //표준공수계(시간)
				oForm.Items.Item("TStdAmt02").Specific.VALUE = TStdAmt;	  //표준공수계(금액)
				oForm.Items.Item("TWkTime02").Specific.VALUE = TWkTime;	  //실동공수계(시간)
				oForm.Items.Item("TWkAmt02").Specific.VALUE = TWkAmt;	  //실동공수계(금액)
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP989_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_PP989_Print_Report01()
		{
			// 이해안됨 원래 에러남
			string WinTitle;
			string ReportName;

			string BPLID;    //사업장
			string TeamCode; //팀
			string RspCode;  //담당
			string ClsCode;  //반
			string CardType; //거래처구분
			string ItemType; //품목구분
			string WCYN;     //생산완료여부
			string DateStd;  //일자기준
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string ItemCode; //품목코드(작번)
			string CpCode;   //공정
			string OrdGbn;   //작업구분
			string CntcCode; //조회자 사번

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID04").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode04").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode04").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode04").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType04").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType04").Specific.Value.ToString().Trim();
				WCYN = oForm.Items.Item("WCYN04").Specific.Value.ToString().Trim();
				DateStd = oForm.Items.Item("DateStd04").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt04").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt04").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode04").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode04").Specific.Value.ToString().Trim();
				OrdGbn = oForm.Items.Item("OrdGbn04").Specific.Value.ToString().Trim();
				CntcCode = dataHelpClass.User_MSTCOD();

				WinTitle = "[PS_PP989] 레포트";
				ReportName = "PS_PP989_01.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@WCYN", WCYN));
				dataPackParameter.Add(new PSH_DataPackClass("@DateStd", DateStd));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdGbn", OrdGbn));
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP989_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_PP989_Print_Report02()
		{
			string WinTitle;
			string ReportName;

			string BPLID;    //사업장
			string TeamCode; //팀
			string RspCode;  //담당
			string ClsCode;  //반
			string CardType; //거래처구분
			string ItemType; //품목구분
			string WCYN;     //생산완료여부
			string DateStd;  //일자기준
			string FrDt;     //기간(Fr)
			string ToDt;     //기간(To)
			string ItemCode; //품목코드(작번)
			string CpCode;   //공정
			string OrdGbn;   //작업구분
			string CntcCode; //조회자 사번

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode02").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode02").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode02").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Value.ToString().Trim();
				WCYN = oForm.Items.Item("WCYN02").Specific.Value.ToString().Trim();
				DateStd = oForm.Items.Item("DateStd02").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode02").Specific.Value.ToString().Trim();
				OrdGbn = oForm.Items.Item("OrdGbn02").Specific.Value.ToString().Trim();
				CntcCode = dataHelpClass.User_MSTCOD();

				WinTitle = "[PS_PP989] 레포트";
				ReportName = "PS_PP989_01.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@WCYN", WCYN));
				dataPackParameter.Add(new PSH_DataPackClass("@DateStd", DateStd));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdGbn", OrdGbn));
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
						PS_PP989_MTX01();
					}
					else if (pVal.ItemUID == "BtnPrt01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP989_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}

					if (pVal.ItemUID == "BtnSrch02")
					{
							PS_PP989_MTX02();
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP989_Print_Report02);
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode02", "");
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
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						PS_PP989_FlushToItemValue(pVal.ItemUID, 0, "");
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row == 0)
						{
							oMat.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat.FlushToDataSource();
						}
						else
						{
							PS_PP989_GetDetail(oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), oMat.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
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
					if (pVal.ItemChanged == true)
					{
						PS_PP989_FlushToItemValue(pVal.ItemUID, 0, "");
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
					PS_PP989_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP989L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP989O);
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							oForm.Freeze(true);
							PS_PP989_Add_MatrixRow01(oMat.VisualRowCount, false);
							oForm.Freeze(false);
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
						case "1285": //복원
							break;
						case "1293": //행삭제
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP989O.RemoveRecord(oDS_PS_PP989O.Size - 1);
							oMat.LoadFromDataSource();
							oForm.Freeze(false);
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_FormDataEvent
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
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
