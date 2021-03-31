using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 보조원장   PS_FI180
	/// </summary>
	internal class PS_FI180 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		
		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI180.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI180_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI180");                 // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);      // 폼 할당
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
                CreateItems();
                ComboBox_Setting();

                oForm.EnableMenu("1283", false);            // 삭제
				oForm.EnableMenu("1286", false);            // 닫기
				oForm.EnableMenu("1287", false);            // 복제
				oForm.EnableMenu("1284", false);            // 취소
				oForm.EnableMenu("1293", false);            // 행삭제
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// CreateItems
		/// </summary>
		private void CreateItems()
	    {
			try
			{
				oForm.DataSources.UserDataSources.Add("StrDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StrDate").Specific.DataBind.SetBound(true, "", "StrDate");
				oForm.DataSources.UserDataSources.Item("StrDate").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");
				oForm.DataSources.UserDataSources.Item("EndDate").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("ChkBox01").Specific.ValOn = "Y";
				oForm.Items.Item("ChkBox01").Specific.ValOff = "N";
				oForm.Items.Item("ChkBox01").Specific.DataBind.SetBound(true, "", "Check01");
				oForm.DataSources.UserDataSources.Item("Check01").Value = "N";    // 미체크로 값을 주고 폼을 로드

				oForm.DataSources.UserDataSources.Add("Check02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("ChkBox02").Specific.ValOn = "Y";
				oForm.Items.Item("ChkBox02").Specific.ValOff = "N";
				oForm.Items.Item("ChkBox02").Specific.DataBind.SetBound(true, "", "Check02");
				oForm.DataSources.UserDataSources.Item("Check02").Value = "N";    // 미체크로 값을 주고 폼을 로드

				//기준일자 콤보_S
				oForm.DataSources.UserDataSources.Add("DateCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DateCls").Specific.DataBind.SetBound(true, "", "DateCls");
				//기준일자 콤보_E
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
            }
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet01.DoQuery(sQry);
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("0", "전체 사업장");
				while (!oRecordSet01.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
					oRecordSet01.MoveNext();
				}

				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 항목선택
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("0", "전체항목");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("1", "관리항목 1");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("2", "관리항목 2");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("3", "관리항목 3");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("4", "관리항목 4");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("5", "관리항목 5");
				oForm.Items.Item("Rpt").Specific.ValidValues.Add("6", "관리항목 6");
				oForm.Items.Item("Rpt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//기준일자 콤보
				oForm.Items.Item("DateCls").Specific.ValidValues.Add("01", "전기일기준");
				oForm.Items.Item("DateCls").Specific.ValidValues.Add("02", "증빙일기준");
				oForm.Items.Item("DateCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("StrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
			string sQry = null;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			//Header
			try
			{
				switch (oUID)
				{
					case "Rpt":
						if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "1")
						{
							sQry = "        Select  TOP 1 ";
							sQry = sQry + "         U_Rpttxt01 ";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt01 IS NOT NULL";
						}
						else if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "2")
						{
							sQry = "        Select  TOP 1";
							sQry = sQry + "         U_Rpttxt02";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt02 IS NOT NULL";
						}
						else if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "3")
						{
							sQry = "        Select  TOP 1 ";
							sQry = sQry + "         U_Rpttxt03 ";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt03 IS NOT NULL";
						}
						else if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "4")
						{
							sQry = "        Select  TOP 1 ";
							sQry = sQry + "         U_Rpttxt04 ";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt04 IS NOT NULL";
						}
						else if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "5")
						{
							sQry = "        Select  TOP 1 ";
							sQry = sQry + "         U_Rpttxt05 ";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt05 IS NOT NULL";
						}
						else if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() == "6")
						{
							sQry = "        Select  TOP 1 ";
							sQry = sQry + "         U_Rpttxt06 ";
							sQry = sQry + " from    [ZMDC_JDT1]";
							sQry = sQry + " where   AcctCode = '" + oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim() + "'";
							sQry = sQry + "         AND U_Rpttxt06 IS NOT NULL";
						}
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("Rpttxt").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}

				// Line
				if (oUID == "Mat01")
				{
					//switch (oCol)
					//{
					//}
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			short ErrNum = 0;
			try
			{
				// Check
				if (string.IsNullOrEmpty(oForm.Items.Item("StrDate").Specific.Value.ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("EndDate").Specific.Value.ToString().Trim()))
				{
					ErrNum = 2;
					throw new Exception();
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("기간 시작일은 필수사항입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("기간 종료일은 필수사항입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				functionReturnValue = false;
			}

			return functionReturnValue;

		}

		/// <summary>
		/// Print_Query
		/// </summary>
		[STAThread]
		private void Print_Query()
		{
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string sQry = string.Empty;

			string StrDate = string.Empty;
			string EndDate = string.Empty;
			string SAcctCode = string.Empty;
			string EAcctCode = string.Empty;
			string BPLID = string.Empty;
			string Rpt = string.Empty;
			string Rpttxt = string.Empty;
			string StrRpt = string.Empty;
			string EndRpt = string.Empty;
			string Summary = string.Empty;
			string Level5 = string.Empty;
			string DateCls = string.Empty;

			StrDate = oForm.Items.Item("StrDate").Specific.Value.ToString().Trim();
			EndDate = oForm.Items.Item("EndDate").Specific.Value.ToString().Trim();
			SAcctCode = oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim();
			EAcctCode = oForm.Items.Item("EAcctCode").Specific.Value.ToString().Trim();
			BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
			Rpt = oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim();
			Rpttxt = oForm.Items.Item("Rpttxt").Specific.Value.ToString().Trim();
			StrRpt = oForm.Items.Item("StrRpt").Specific.Value.ToString().Trim();
			EndRpt = oForm.Items.Item("EndRpt").Specific.Value.ToString().Trim();
			Summary = oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim();
			Level5 = oForm.DataSources.UserDataSources.Item("Check02").Value.ToString().Trim();
			DateCls = oForm.Items.Item("DateCls").Specific.Selected.Value.ToString().Trim();

			if (string.IsNullOrEmpty(SAcctCode))
			{
				SAcctCode = "1";
			}
			if (string.IsNullOrEmpty(EAcctCode))
			{
				EAcctCode = "9999999999";
			}
			if (string.IsNullOrEmpty(StrRpt))
			{
				StrRpt = "!";
			}
			if (string.IsNullOrEmpty(EndRpt))
			{
				EndRpt = "ZZZZZZZZZZ";
			}

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				WinTitle = "[PS_FI180] 보조원장";
				if (Rpt == "0")
				{
					ReportName = "PS_FI180_00.RPT";
				}
				else if (Rpt == "1")
				{
					ReportName = "PS_FI180_01.RPT";
				}
				else if (Rpt == "2")
				{
					ReportName = "PS_FI180_02.RPT";
				}
				else if (Rpt == "3")
				{
					ReportName = "PS_FI180_03.RPT";
				}
				else if (Rpt == "4")
				{
					ReportName = "PS_FI180_04.RPT";
				}
				else if (Rpt == "5")
				{
					ReportName = "PS_FI180_05.RPT";
				}
				else if (Rpt == "6")
				{
					ReportName = "PS_FI180_06.RPT";
				}

				if (Summary == "Y")
				{
					WinTitle = "[PS_FI180] 보조원장 집계표";
					ReportName = "PS_FI180_20.RPT";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula
				dataPackFormula.Add(new PSH_DataPackClass("@StrDate", StrDate.Substring(0, 4) + "-" + StrDate.Substring(4, 2) + "-" + StrDate.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@EndDate", EndDate.Substring(0, 4) + "-" + EndDate.Substring(4, 2) + "-" + EndDate.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackFormula.Add(new PSH_DataPackClass("@Rpt", Rpt)); // 출력구분


				//System.DateTime RpmtDate = default(System.DateTime);  //변수
				//RpmtDate = DateTime.ParseExact(oForm.Items.Item("RpmtDate").Specific.Value, "yyyyMMdd", null);  //인자MOVE


				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StrDate", DateTime.ParseExact(StrDate, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@EndDate", DateTime.ParseExact(EndDate, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@SAcctCode", SAcctCode));
				dataPackParameter.Add(new PSH_DataPackClass("@EAcctCode", EAcctCode));
				dataPackParameter.Add(new PSH_DataPackClass("@Rpt", Rpt));
				dataPackParameter.Add(new PSH_DataPackClass("@Rpttxt", Rpttxt));
				dataPackParameter.Add(new PSH_DataPackClass("@StrRpt", StrRpt));
				dataPackParameter.Add(new PSH_DataPackClass("@EndRpt", EndRpt));
				dataPackParameter.Add(new PSH_DataPackClass("@Summary", Summary));
				dataPackParameter.Add(new PSH_DataPackClass("@Level5", Level5));
				dataPackParameter.Add(new PSH_DataPackClass("@DateCls", DateCls));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")   // 출력버튼 클릭시
                    {
                        if (HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(Print_Query);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
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
                        // 헤더
                        if (pVal.ItemUID == "SAcctCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SAcctCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "EAcctCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("EAcctCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }

                        if (pVal.ItemUID == "StrRpt")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("StrRpt").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "EndRpt")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("EndRpt").Specific.Value))
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// COMBO_SELECT 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
                    if (pVal.ItemUID == "Rpt")
                    {
                        if (oForm.Items.Item("Rpt").Specific.Selected.Value.ToString().Trim() != "0")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim()))
                            {
                                oForm.Freeze(true);
                                PSH_Globals.SBO_Application.MessageBox("계정과목(시작) 입력 후 항목선택을 하여 주시기 바랍니다.");
                                oForm.Items.Item("Rpt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("SAcctCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                                BubbleEvent = false;
                                return;
                            }
                            else if (!string.IsNullOrEmpty(oForm.Items.Item("SAcctCode").Specific.Value.ToString().Trim()))
                            {
                                oForm.Freeze(true);
                                FlushToItemValue(pVal.ItemUID, 0, "");
                                oForm.Items.Item("StrRpt").Enabled = true;
                                oForm.Items.Item("EndRpt").Enabled = true;
                                oForm.Items.Item("EndRpt").Specific.Value = "";
                                oForm.Items.Item("StrRpt").Specific.Value = "";
                                oForm.Freeze(false);
                            }
                        }
                        else
                        {
                            oForm.Freeze(true);
                            oForm.Items.Item("Rpttxt").Specific.Value = "";
                            oForm.Items.Item("StrRpt").Specific.Value = "";
                            oForm.Items.Item("EndRpt").Specific.Value = "";
                            oForm.Items.Item("StrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("StrRpt").Enabled = false;
                            oForm.Items.Item("EndRpt").Enabled = false;
                            oForm.Freeze(false);
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
				oForm.Freeze(false);
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
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "SAcctCode")
                        {
                            oForm.Items.Item("Rpt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("Rpttxt").Specific.Value = "";
                        }
                    }
                }
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				BubbleEvent = false;
			}
			finally
			{
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
					SubMain.Remove_Forms(oFormUniqueID01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_MenuEvent
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
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1285":							//복원
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
					}
			    // BeforeAction = False
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1285":							//복원
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
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
				// BeforeAction = True
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
							////33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
							////34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
							////35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
							////36
							break;
					}
					// BeforeAction = False
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
							////33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
							////34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
							////35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
							////36
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
			}
		}

		/// <summary>
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			// ERROR: Not supported in C#: OnErrorStatement
			try
			{
				if ((eventInfo.BeforeAction == true))
				{

				}
				else if ((eventInfo.BeforeAction == false))
				{
					// 작업
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}
	}
}
