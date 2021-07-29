using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 바코드 리스트 출력
	/// </summary>
	internal class PS_PP550 : PSH_BaseClass
	{
		private string oFormUniqueID;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP550.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP550_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP550");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP550_CreateItems();
				PS_PP550_SetComboBox();
				PS_PP550_EnableFormItem();
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
		/// PS_PP550_CreateItems
		/// </summary>
		private void PS_PP550_CreateItems()
		{
			try
			{
				//사원바코드
				//사업장1
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//팀1
				oForm.DataSources.UserDataSources.Add("TeamCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TeamCode01").Specific.DataBind.SetBound(true, "", "TeamCode01");

				//담당1
				oForm.DataSources.UserDataSources.Add("RspCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("RspCode01").Specific.DataBind.SetBound(true, "", "RspCode01");

				//반1
				oForm.DataSources.UserDataSources.Add("ClsCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ClsCode01").Specific.DataBind.SetBound(true, "", "ClsCode01");

				//구분1
				oForm.DataSources.UserDataSources.Add("JIGTYP01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("JIGTYP01").Specific.DataBind.SetBound(true, "", "JIGTYP01");

				//작번바코드
				//사업장2
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//거래처구분2
				oForm.DataSources.UserDataSources.Add("CardType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("CardType02").Specific.DataBind.SetBound(true, "", "CardType02");

				//품목구분2
				oForm.DataSources.UserDataSources.Add("ItemType02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItemType02").Specific.DataBind.SetBound(true, "", "ItemType02");

				//기준년월2
				oForm.DataSources.UserDataSources.Add("StdYM02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("StdYM02").Specific.DataBind.SetBound(true, "", "StdYM02");

				oForm.Items.Item("StdYM02").Specific.Value = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP550_SetComboBox
		/// </summary>
		private void PS_PP550_SetComboBox()
		{
			string BPLID;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//사업장1
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

				//직원구분
				sQry = " SELECT      '%' AS [Code],";
				sQry += "                '전체' AS [Name]";
				sQry += " UNION ALL";
				sQry += " SELECT      U_Code AS [Code],";
				sQry += "                U_CodeNm";
				sQry += " FROM       [@PS_HR200L]";
				sQry += " WHERE      Code = 'P126'";
				sQry += "                AND U_UseYN= 'Y'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("JIGTYP01").Specific, sQry, "%", false, false);

				//사업장2
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//거래처구분2
				sQry = " SELECT      '%' AS [Code],";
				sQry += "                '전체' AS [Name]";
				sQry += " UNION ALL";
				sQry += " SELECT     T0.U_Minor AS [Code],";
				sQry += "               T0.U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L] AS T0";
				sQry += " WHERE     T0.Code = 'C100'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType02").Specific, sQry, "%", false, false);

				//품목구분2
				sQry = " SELECT      '%' AS [Code],";
				sQry += "                '전체' AS [Name]";
				sQry += " UNION ALL";
				sQry += " SELECT     T0.U_Minor AS [Code],";
				sQry += "               T0.U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L] AS T0";
				sQry += " WHERE     T0.Code = 'S002'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType02").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP550_EnableFormItem
		/// </summary>
		private void PS_PP550_EnableFormItem()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLID01").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
					PS_PP550_FlushToItemValue("BPLID01", 0, "");  //팀, 담당, 반 콤보박스 강제 설정
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
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP550_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;
			string BPLID;
			string TeamCode;
			string RspCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "BPLID01":
						BPLID = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("TeamCode01").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode01").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode01").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//부서콤보세팅
						oForm.Items.Item("TeamCode01").Specific.ValidValues.Add("%", "전체");
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '1'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char2 = '" + BPLID + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode01").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

						if (oForm.Items.Item("RspCode01").Specific.ValidValues.Count == 0)
						{
							oForm.Items.Item("RspCode01").Specific.ValidValues.Add("%", "전체");
							oForm.Items.Item("RspCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("ClsCode01").Specific.ValidValues.Add("%", "전체");
							oForm.Items.Item("ClsCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						}
						break;

					case "TeamCode01":
						TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("RspCode01").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("RspCode01").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("RspCode01").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//담당콤보세팅
						oForm.Items.Item("RspCode01").Specific.ValidValues.Add("%", "전체");
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '2'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode01").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "RspCode01":
						TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim();
						RspCode  = oForm.Items.Item("RspCode01").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("ClsCode01").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("ClsCode01").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("ClsCode01").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//반콤보세팅
						oForm.Items.Item("ClsCode01").Specific.ValidValues.Add("%", "전체");
						sQry = " SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '9'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + RspCode + "'";
						sQry += "                 AND U_Char2 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode01").Specific, sQry, "", false, false);
						oForm.Items.Item("ClsCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP550_PrintReport01
		/// </summary>
		[STAThread]
		private void PS_PP550_PrintReport01()
		{
			string WinTitle;
			string ReportName;
			string BPLID;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;	 //반
			string JIGTYP;   //직급구분
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim();
				RspCode  = oForm.Items.Item("RspCode01").Specific.Value.ToString().Trim();
				ClsCode  = oForm.Items.Item("ClsCode01").Specific.Value.ToString().Trim();
				JIGTYP   = oForm.Items.Item("JIGTYP01").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP550] 레포트";
				ReportName = "PS_PP550_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
				dataPackParameter.Add(new PSH_DataPackClass("@JIGTYP", JIGTYP));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP550_PrintReport02
		/// </summary>
		[STAThread]
		private void PS_PP550_PrintReport02()
		{
			string WinTitle;
			string ReportName;
			string BPLID;   
			string CardType;
			string ItemType;
			string StdYM;   //기준년월
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID    = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType02").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType02").Specific.Value.ToString().Trim();
				StdYM    = oForm.Items.Item("StdYM02").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP550] 레포트";
				ReportName = "PS_PP550_02.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

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
					if (pVal.ItemUID == "BtnPrt01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP550_PrintReport01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP550_PrintReport02);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnPrt01";
					}

					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnPrt02";
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
					PS_PP550_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_PP550_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
