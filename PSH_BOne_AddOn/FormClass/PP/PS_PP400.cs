using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 납기초과 생산품 현황
	/// </summary>
	internal class PS_PP400 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.DataTable oDS_PS_PP400A;
		private SAPbouiCOM.DataTable oDS_PS_PP400B;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP400.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP400_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP400");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP400_CreateItems();
				PS_PP400_SetComboBox();
				PS_PP400_Initialize();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
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
		/// PS_PP400_CreateItems
		/// </summary>
		private void PS_PP400_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;

				oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oGrid02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

				oForm.DataSources.DataTables.Add("PS_PP400A");
				oForm.DataSources.DataTables.Add("PS_PP400B");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP400A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP400B");

				oDS_PS_PP400A = oForm.DataSources.DataTables.Item("PS_PP400A");
				oDS_PS_PP400B = oForm.DataSources.DataTables.Item("PS_PP400B");

				//납기도래품목 조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

				//(작번)품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode01").Specific.DataBind.SetBound(true, "", "ItemCode01");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName01").Specific.DataBind.SetBound(true, "", "ItemName01");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemSpec01").Specific.DataBind.SetBound(true, "", "ItemSpec01");

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardCode01").Specific.DataBind.SetBound(true, "", "CardCode01");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CardName01").Specific.DataBind.SetBound(true, "", "CardName01");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("CardType01").Specific.DataBind.SetBound(true, "", "CardType01");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItemType01").Specific.DataBind.SetBound(true, "", "ItemType01");

				//생산구분
				oForm.DataSources.UserDataSources.Add("TradType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TradType01").Specific.DataBind.SetBound(true, "", "TradType01");

				//연간품여부
				oForm.DataSources.UserDataSources.Add("YearPdYN01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("YearPdYN01").Specific.DataBind.SetBound(true, "", "YearPdYN01");

				//장비/공구
				oForm.DataSources.UserDataSources.Add("ItemCls01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItemCls01").Specific.DataBind.SetBound(true, "", "ItemCls01");

				//자체/외주
				oForm.DataSources.UserDataSources.Add("InOut01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("InOut01").Specific.DataBind.SetBound(true, "", "InOut01");

				//D-Day
				oForm.DataSources.UserDataSources.Add("DMinuDay01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("DMinuDay01").Specific.DataBind.SetBound(true, "", "DMinuDay01");

				//D+Day
				oForm.DataSources.UserDataSources.Add("DPlusDay01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("DPlusDay01").Specific.DataBind.SetBound(true, "", "DPlusDay01");

				//영업납품완료 제외
				oForm.DataSources.UserDataSources.Add("SaleCplt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("SaleCplt01").Specific.DataBind.SetBound(true, "", "SaleCplt01");

				//납기초과생산품 조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");

				//(작번)품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode02").Specific.DataBind.SetBound(true, "", "ItemCode02");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName02").Specific.DataBind.SetBound(true, "", "ItemName02");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemSpec02").Specific.DataBind.SetBound(true, "", "ItemSpec02");

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardCode02").Specific.DataBind.SetBound(true, "", "CardCode02");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CardName02").Specific.DataBind.SetBound(true, "", "CardName02");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP400_SetComboBox
		/// </summary>
		private void PS_PP400_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);

				//거래처구분
				sQry = "  SELECT  U_Minor AS [Code], ";
				sQry += "         U_CdName AS [Name]";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'C100'";
				oForm.Items.Item("CardType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType01").Specific, sQry, "%", false, false);

				//품목구분
				sQry = "  SELECT  U_Minor AS [Code], ";
				sQry += "         U_CdName AS [Name]";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'S002'";
				oForm.Items.Item("ItemType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType01").Specific, sQry, "%", false, false);

				//생산구분
				oForm.Items.Item("TradType01").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("TradType01").Specific.ValidValues.Add("1", "일반");
				oForm.Items.Item("TradType01").Specific.ValidValues.Add("3", "선생산");
				oForm.Items.Item("TradType01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//연간품여부
				oForm.Items.Item("YearPdYN01").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("YearPdYN01").Specific.ValidValues.Add("Y", "Y");
				oForm.Items.Item("YearPdYN01").Specific.ValidValues.Add("N", "N");
				oForm.Items.Item("YearPdYN01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//장비/공구
				oForm.Items.Item("ItemCls01").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("ItemCls01").Specific.ValidValues.Add("T", "공구");
				oForm.Items.Item("ItemCls01").Specific.ValidValues.Add("M", "장비");
				oForm.Items.Item("ItemCls01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//자체/외주
				oForm.Items.Item("InOut01").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InOut01").Specific.ValidValues.Add("IN", "자체");
				oForm.Items.Item("InOut01").Specific.ValidValues.Add("OUT", "외주");
				oForm.Items.Item("InOut01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//납기초과생산품 조회
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP400_Initialize
		/// </summary>
		private void PS_PP400_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//아이디별 사업장 세팅
				oForm.Items.Item("BPLID01").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("BPLID02").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.ToString("yyyy") + "0101";
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("DMinuDay01").Specific.Value = "-10";
				oForm.Items.Item("DPlusDay01").Specific.Value = "10";

				oForm.Items.Item("Folder01").Specific.Select();  //폼이 로드 될 때 Folder01이 선택됨
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP400_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP400_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "ItemCode01":
						oForm.Items.Item("ItemName01").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("ItemCode01").Specific.Value.ToString().Trim() + "'", "");	//작번
						oForm.Items.Item("ItemSpec01").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("ItemCode01").Specific.Value.ToString().Trim() + "'", ""); //규격
						break;
					case "CardCode01":
						oForm.Items.Item("CardName01").Specific.Value = dataHelpClass.Get_ReData("CardName", "CardCode", "OCRD", "'" + oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim() + "'", "");	//거래처
						break;
					case "ItemCode02":
						oForm.Items.Item("ItemName02").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim() + "'", "");  //작번
						oForm.Items.Item("ItemSpec02").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim() + "'", "");  //규격
						break;
					case "CardCode02":
						oForm.Items.Item("CardName02").Specific.Value = dataHelpClass.Get_ReData("CardName", "CardCode", "OCRD", "'" + oForm.Items.Item("CardCode02").Specific.Value.ToString().Trim() + "'", "");  //거래처
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP400_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP400_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					errMessage = "사업장은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim()))
				{
					errMessage = "일자를 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목분류코드를 확인하여 주십시오.";
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
		/// PS_PP400_MTX01
		/// </summary>
		private void PS_PP400_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLId;	  //사업장
			string CardCode;
			string ItemCode;
			string FrDt;	  //기간(Fr)
			string ToDt;	  //기간(To)
			string CardType; //거래처구분
			string ItemType; //품목구분
			string TradType; //생산구분
			string YearPdYN; //연간품여부
			string ItemCls;	  //장비/공구
			string InOut;	  //자체/외주
			string DMinusDay; //D-일수
			string DPlusDay;  //D+일수
			string SaleCplt;  //영업납품완료 제외여부

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				CardCode  = oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim();
				ItemCode  = oForm.Items.Item("ItemCode01").Specific.Value.ToString().Trim();
				FrDt      = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt      = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				CardType  = oForm.Items.Item("CardType01").Specific.Selected.Value.ToString().Trim();
				ItemType  = oForm.Items.Item("ItemType01").Specific.Selected.Value.ToString().Trim();
				TradType  = oForm.Items.Item("TradType01").Specific.Selected.Value.ToString().Trim();
				YearPdYN  = oForm.Items.Item("YearPdYN01").Specific.Selected.Value.ToString().Trim();
				ItemCls   = oForm.Items.Item("ItemCls01").Specific.Selected.Value.ToString().Trim();
				InOut     = oForm.Items.Item("InOut01").Specific.Selected.Value.ToString().Trim();
				DMinusDay = oForm.Items.Item("DMinuDay01").Specific.Value.ToString().Trim();
				DPlusDay  = oForm.Items.Item("DPlusDay01").Specific.Value.ToString().Trim();

				if (oForm.Items.Item("SaleCplt01").Specific.Checked == true)
				{
					SaleCplt = "Y";
				}	
				else
				{
					SaleCplt = "N";
				}

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC PS_PP400_01 '";
				sQry += BPLId + "','";
				sQry += CardCode + "','";
				sQry += ItemCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += TradType + "','";
				sQry += YearPdYN + "','";
				sQry += ItemCls + "','";
				sQry += InOut + "','";
				sQry += DMinusDay + "','";
				sQry += DPlusDay + "','";
				sQry += SaleCplt + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP400A.ExecuteQuery(sQry);

				oGrid01.Columns.Item(7).RightJustified = true;
				oGrid01.Columns.Item(8).RightJustified = true;
				oGrid01.Columns.Item(9).RightJustified = true;
				oGrid01.Columns.Item(13).RightJustified = true;
				oGrid01.Columns.Item(18).RightJustified = true;
				oGrid01.Columns.Item(20).RightJustified = true;
				oGrid01.Columns.Item(21).RightJustified = true;
				oGrid01.Columns.Item(25).RightJustified = true;
				oGrid01.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //잔여납기일, 노랑

				if (oGrid01.Rows.Count == 0)
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
		/// PS_PP400_MTX02
		/// </summary>
		private void PS_PP400_MTX02()
		{
			string sQry;
			string errMessage = string.Empty;
			string BPLId;     //사업장
			string FrDt;      //기간(Fr)
			string ToDt;      //기간(To)
			string ItemCode;  //공정
			string CardCode;  //작업구분

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLID02").Specific.Selected.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode02").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode02").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP400_02] '";
				sQry += BPLId + "','";
				sQry += CardCode + "','";
				sQry += ItemCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP400B.ExecuteQuery(sQry);

				oGrid02.Columns.Item(8).RightJustified = true;
				oGrid02.Columns.Item(9).RightJustified = true;
				oGrid02.Columns.Item(10).RightJustified = true;
				oGrid02.Columns.Item(11).RightJustified = true;
				oGrid02.Columns.Item(14).RightJustified = true;

				if (oGrid02.Rows.Count == 0)
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
		/// PS_PP400_PrintReport01
		/// </summary>
		[STAThread]
		private void PS_PP400_PrintReport01()
		{
			string WinTitle;
			string ReportName;
			string BPLId;
			string FrDt;
			string ToDt;
			string CardType;
			string ItemCode;
			string CardCode;
			string ItemType;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId    = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode01").Specific.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType01").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType01").Specific.Selected.Value.ToString().Trim();

				WinTitle = "[PS_PP400_01] 납기도래품목 조회";
				ReportName = "PS_PP400_01.RPT";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP400_PrintReport02
		/// </summary>
		[STAThread]
		private void PS_PP400_PrintReport02()
		{
			string WinTitle;
			string ReportName;

			string BPLId;
			string DocDateFr;
			string DocDateTo;
			string ItemCode;
			string CardCode;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode01").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode01").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("FrDt02").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("ToDt02").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP400_02] 납기초과생산품 조회";
				ReportName = "PS_PP400_02.RPT";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2)));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
						PS_PP400_MTX01();
					}
					else if (pVal.ItemUID == "BtnSrch02")  //납기초과생산품 조회
					{
						PS_PP400_MTX02();
					}
					else if (pVal.ItemUID == "BtnPrt01")  //납기초과생산품 조회(리포트)
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP400_PrintReport01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt02")  //납기초과생산품 조회(리포트)
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP400_PrintReport02);
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
					else if (pVal.ItemUID == "Folder02")
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode02", "");
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
						PS_PP400_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP400A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP400B);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
