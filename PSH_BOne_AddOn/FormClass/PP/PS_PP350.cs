using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산일보
	/// </summary>
	internal class PS_PP350 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private short p_prt;
		private SAPbouiCOM.Grid oGrid;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP350.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP350_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP350");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP350_CreateItems();
				PS_PP350_SetComboBox();
				PS_PP350_Initialize();

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
		/// PS_PP350_CreateItems
		/// </summary>
		private void PS_PP350_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				oForm.DataSources.DataTables.Add("PS_PP350");

				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("일자", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("문서번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("작지번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("품목코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("품목명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("공정코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("공정명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("작업자명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("배치번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("생산수량", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("합격수량", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("불량수량", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("스크랩", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("LOSS", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("작업시간", SAPbouiCOM.BoFieldsType.ft_Float);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("원재료코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
				oForm.DataSources.DataTables.Item("PS_PP350").Columns.Add("원재료명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP350");

				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP350_SetComboBox
		/// </summary>
		private void PS_PP350_SetComboBox()
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

				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "%", "전체");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "10", "일반");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "20", "PSMT지원");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "30", "외주");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "40", "실적");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "50", "일반조정");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "60", "외주조정");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP350", "OrdType", "", "70", "설계시간");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP350", "OrdType", false);
				oForm.Items.Item("OrdType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_PP350_Initialize
		/// </summary>
		private void PS_PP350_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{	
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP350_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP350_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "ItmBsort":
						sQry = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code =  '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItmBName").Specific.String = oRecordSet.Fields.Item("Name").Value.ToString().Trim();
						break;
					case "ItemCode":
						sQry = "SELECT ItemName FROM [OITM] WHERE ItemCode =  '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItemName").Specific.String = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
						break;
					case "CpCode":
						sQry = "SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode =  '" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CpName").Specific.String = oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim();
						break;
					case "CItemCod":
						sQry = "SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "' and U_ItemCod2 = '" + oForm.Items.Item("CItemCod").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CItemNam").Specific.String = oRecordSet.Fields.Item("U_ItemNam2").Value.ToString().Trim();
						break;
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
		/// PS_PP350_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP350_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
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
		/// PS_PP350_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP350_PrintReport()
		{
			string sQry;
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string BPLName;
			string BPLId;
			string DocDateFr;
			string DocDateTo;
			string ItmBsort;
			string ItemCode;
			string CpCode;
			string OrdType;
			string CItemCod;
			string OrdNum;
			string BatchNum;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				CItemCod = oForm.Items.Item("CItemCod").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				BatchNum = oForm.Items.Item("BatchNum").Specific.Value.ToString().Trim();

				sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + BPLId + "'";
				oRecordSet.DoQuery(sQry);
				BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				if (p_prt == 1)
				{
					WinTitle = "생산일보일자별 [PS_PP350_01]";
					ReportName = "PS_PP350_01.RPT";
				}
				else if (p_prt == 2)
				{
					WinTitle = "생산일보공정별 [PS_PP350_03]";
					ReportName = "PS_PP350_03.RPT";
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLName));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdGbn", ItmBsort));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdType", OrdType));
				dataPackParameter.Add(new PSH_DataPackClass("@CItemCod", CItemCod));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
				dataPackParameter.Add(new PSH_DataPackClass("@BatchNum", BatchNum));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP350_MTX01
		/// </summary>
		private void PS_PP350_MTX01()
		{
			int i;
			string sQry;
			string[] COLNAM = new string[17];
			string errMessage = string.Empty;
			string BPLId;
			string DocDateFr;
			string DocDateTo;
			string ItmBsort;
			string ItemCode;
			string CpCode;
			string OrdType;
			string CItemCod;
			string OrdNum;
			string BatchNum;

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				CItemCod = oForm.Items.Item("CItemCod").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				BatchNum = oForm.Items.Item("BatchNum").Specific.Value.ToString().Trim();

				sQry = "EXEC [PS_PP350_02] '" + BPLId + "', '" + ItmBsort + "', '" + DocDateFr + "', '" + DocDateTo + "', '" + ItemCode + "', '" + CpCode + "', '" + OrdType + "', '" + CItemCod + "', '" + OrdNum + "', '" + BatchNum + "'";

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				COLNAM[0] = "일자";
				COLNAM[1] = "문서번호";
				COLNAM[2] = "작지번호";
				COLNAM[3] = "품목코드";
				COLNAM[4] = "품목명";
				COLNAM[5] = "공정코드";
				COLNAM[6] = "공정명";
				COLNAM[7] = "작업자명";
				COLNAM[8] = "배치번호";
				COLNAM[9] = "생산수량";
				COLNAM[10] = "합격수량";
				COLNAM[11] = "불량수량";
				COLNAM[12] = "스크랩";
				COLNAM[13] = "LOSS";
				COLNAM[14] = "작업시간";
				COLNAM[15] = "원재료코드";
				COLNAM[16] = "원재료명";

				for (i = 0; i < COLNAM.Length; i++)
				{
					oGrid.Columns.Item(i).TitleObject.Caption = COLNAM[i];
				}

				oGrid.Columns.Item(9).RightJustified = true;
				oGrid.Columns.Item(10).RightJustified = true;
				oGrid.Columns.Item(11).RightJustified = true;
				oGrid.Columns.Item(12).RightJustified = true;
				oGrid.Columns.Item(13).RightJustified = true;

				if (oGrid.Rows.Count == 0)
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
					if (pVal.ItemUID == "1")
					{
					}
					else if (pVal.ItemUID == "BtnPrint" || pVal.ItemUID == "BtnPrint1")
					{
						if (PS_PP350_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							if (pVal.ItemUID == "BtnPrint")
							{
								p_prt = 1;
							}
							else
							{
								p_prt = 2;
							}

							System.Threading.Thread thread = new System.Threading.Thread(PS_PP350_PrintReport);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						if (PS_PP350_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							PS_PP350_MTX01();
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmBsort", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CItemCod", "");
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
						if (pVal.ItemUID == "ItmBsort" || pVal.ItemUID == "ItemCode" || pVal.ItemUID == "CpCode" || pVal.ItemUID == "CItemCod")
						{
							PS_PP350_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
