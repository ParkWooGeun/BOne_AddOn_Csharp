using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번이력관리
	/// </summary>
	internal class PS_SD956 : PSH_BaseClass
	{
		private string oFormUniqueID;
		public SAPbouiCOM.Grid oGrid01;
		public SAPbouiCOM.Grid oGrid02;
		public SAPbouiCOM.Grid oGrid03;
		public SAPbouiCOM.Grid oGrid04;

		public SAPbouiCOM.DataTable oDS_PS_SD956A;
		public SAPbouiCOM.DataTable oDS_PS_SD956B;
		public SAPbouiCOM.DataTable oDS_PS_SD956C;
		public SAPbouiCOM.DataTable oDS_PS_SD956D;
			
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD956.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

				oFormUniqueID = "PS_SD956_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD956");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);

                PS_SD956_CreateItems();
                PS_SD956_ComboBox_Setting();
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
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
			}
		}

		/// <summary>
		/// PS_SD956_CreateItems
		/// </summary>
		private void PS_SD956_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid04 = oForm.Items.Item("Grid04").Specific;

				oForm.DataSources.DataTables.Add("PS_SD956A");
				oForm.DataSources.DataTables.Add("PS_SD956B");
				oForm.DataSources.DataTables.Add("PS_SD956C");
				oForm.DataSources.DataTables.Add("PS_SD956D");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_SD956A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_SD956B");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_SD956C");
				oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_SD956D");

				oDS_PS_SD956A = oForm.DataSources.DataTables.Item("PS_SD956A");
				oDS_PS_SD956B = oForm.DataSources.DataTables.Item("PS_SD956B");
				oDS_PS_SD956C = oForm.DataSources.DataTables.Item("PS_SD956C");
				oDS_PS_SD956D = oForm.DataSources.DataTables.Item("PS_SD956D");

				//수주일자(Fr)
				oForm.DataSources.UserDataSources.Add("OrdFrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("OrdFrDt").Specific.DataBind.SetBound(true, "", "OrdFrDt");

				//수주일자(To)
				oForm.DataSources.UserDataSources.Add("OrdToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("OrdToDt").Specific.DataBind.SetBound(true, "", "OrdToDt");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//팀(거래처)
				oForm.DataSources.UserDataSources.Add("CardTeam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CardTeam").Specific.DataBind.SetBound(true, "", "CardTeam");

				//작번
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

				//장비/공구
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//순환품
				oForm.DataSources.UserDataSources.Add("YearPdYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("YearPdYN").Specific.DataBind.SetBound(true, "", "YearPdYN");
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
		/// PS_SD956_ComboBox_Setting
		/// </summary>
		private void PS_SD956_ComboBox_Setting()
		{
			string sQry;
			string BPLID;
			
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				BPLID = dataHelpClass.User_BPLID();

				//거래처구분
				sQry = "  SELECT   U_Minor,";
				sQry += "          U_CdName";
				sQry += " FROM     [@PS_SY001L]";
				sQry += " WHERE    Code = 'C100'";
				sQry += " ORDER BY Code";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래처팀
				sQry = "  SELECT      U_Minor AS [Code],";
				sQry += "             U_CdName As [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'I001'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Minor";
				oForm.Items.Item("CardTeam").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardTeam").Specific, sQry, "", false, false);
				oForm.Items.Item("CardTeam").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//장비/공구
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("M", "장비");
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("T", "공구");
				oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//순환품
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("Y", "Y");
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("N", "N");
				oForm.Items.Item("YearPdYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
		/// Form Item Event
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">pVal</param>
		/// <param name="BubbleEvent">Bubble Event</param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			switch (pVal.EventType) {
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:				//1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:					//2
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:				//5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK:					     //6
					//Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:				//7
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:		//8
					//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE:					//10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:				//11
					//Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:				//18
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:			//19
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:				//20
					Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:			//27
					//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:					//3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:					//4
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:				//17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSrch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SD956_MTX01();							//영업정보
							PS_SD956_MTX02();							//생산정보
							PS_SD956_MTX03();							//구매정보
							PS_SD956_MTX04();							//회계정보
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{

					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
					}
					if (pVal.ItemUID == "Folder04")
					{
						oForm.PaneLevel = 4;
					}
					if (pVal.ItemUID == "PS_SD956")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
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
					PS_SD956_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					if (pVal.ItemUID == "Grid02")
					{
						if (pVal.Row == -1)
						{
							oGrid02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid02.Rows.SelectedRows.Count > 0)  //생산정보
							{
								PS_SD956_GetProductionDetail();
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
							if (oGrid03.Rows.SelectedRows.Count > 0)  //구매정보
							{
								PS_SD956_GetPurchaseDetail();
							}
							else
							{
								BubbleEvent = false;
							}
						}
					}
					else if (pVal.ItemUID == "Grid04")
					{
						if (pVal.Row == -1)
						{
							oGrid04.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						}
						else
						{
							if (oGrid04.Rows.SelectedRows.Count > 0)  //회계정보
							{
								PS_SD956_GetAccountingDetail();
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
			finally
			{
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
					PS_SD956_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD956_FormResize();
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// 
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD956A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD956B);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD956C);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD956D);
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
						case "1284":                        //취소
							break;
						case "1286":                        //닫기
							break;
						case "1293":                        //행삭제
							break;
						case "1281":                        //찾기
							break;
						case "1282":                        //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                        //레코드이동버튼
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":                        //취소
							break;
						case "1286":                        //닫기
							break;
						case "1293":                        //행삭제
							break;
						case "1281":                        //찾기
							break;
						case "1282":                        //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                        //레코드이동버튼
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
				if (BusinessObjectInfo.BeforeAction == true)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
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
		/// PS_SD956_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD956_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "ItemCode":
						oForm.Items.Item("ItemName").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.Items.Item("ItemSpec").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "");
						break;
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
		/// PS_SD956_MTX01  영업정보 조회
		/// </summary>
		private void PS_SD956_MTX01()
		{
			int ErrNum = 0;
			string sQry;

			string OrdFrDt;			//수주기간(Fr)
			string OrdToDt;			//수주기간(To)
			string CardType;		//거래처구분
			string CardTeam;		//팀(거래처)
			string ItemCode;		//작번
			string ItemType;		//품목구분(장비/공구)
			string YearPdYN;        //순환품(연간품)

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdFrDt = oForm.Items.Item("OrdFrDt").Specific.Value.ToString().Trim();
				OrdToDt  = oForm.Items.Item("OrdToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				CardTeam = oForm.Items.Item("CardTeam").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				YearPdYN = oForm.Items.Item("YearPdYN").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_SD956_01 '";
				sQry += OrdFrDt + "','";
				sQry += OrdToDt + "','";
				sQry += CardType + "','";
				sQry += CardTeam + "','";
				sQry += ItemCode + "','";
				sQry += ItemType + "','";
				sQry += YearPdYN + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_SD956A.ExecuteQuery(sQry);

				oGrid01.Columns.Item(5).RightJustified = true;
				oGrid01.Columns.Item(6).RightJustified = true;
				oGrid01.Columns.Item(12).RightJustified = true;
				oGrid01.Columns.Item(14).RightJustified = true;
				oGrid01.Columns.Item(16).RightJustified = true;
				oGrid01.Columns.Item(17).RightJustified = true;
				oGrid01.Columns.Item(18).RightJustified = true;

				if (oGrid01.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
			}

			catch (Exception ex)
			{
				if ( ErrNum == 1)
                {
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
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
		/// PS_SD956_MTX02  생산정보 조회
		/// </summary>
		private void PS_SD956_MTX02()
		{
			int ErrNum = 0;
			string sQry;

			string OrdFrDt;         //수주기간(Fr)
			string OrdToDt;         //수주기간(To)
			string CardType;        //거래처구분
			string CardTeam;        //팀(거래처)
			string ItemCode;        //작번
			string ItemType;        //품목구분(장비/공구)
			string YearPdYN;        //순환품(연간품)

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdFrDt = oForm.Items.Item("OrdFrDt").Specific.Value.ToString().Trim();
				OrdToDt = oForm.Items.Item("OrdToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				CardTeam = oForm.Items.Item("CardTeam").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				YearPdYN = oForm.Items.Item("YearPdYN").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "       EXEC PS_SD956_02 '";
				sQry += OrdFrDt + "','";
				sQry += OrdToDt + "','";
				sQry += CardType + "','";
				sQry += CardTeam + "','";
				sQry += ItemCode + "','";
				sQry += ItemType + "','";
				sQry += YearPdYN + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_SD956B.ExecuteQuery(sQry);

				oGrid02.Columns.Item(10).RightJustified = true;
				oGrid02.Columns.Item(14).RightJustified = true;

				if (oGrid02.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD956_MTX03  구매정보 조회
		/// </summary>
		private void PS_SD956_MTX03()
		{
			int ErrNum = 0;
			string sQry;

			string OrdFrDt;         //수주기간(Fr)
			string OrdToDt;         //수주기간(To)
			string CardType;        //거래처구분
			string CardTeam;        //팀(거래처)
			string ItemCode;        //작번
			string ItemType;        //품목구분(장비/공구)
			string YearPdYN;        //순환품(연간품)

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdFrDt = oForm.Items.Item("OrdFrDt").Specific.Value.ToString().Trim();
				OrdToDt = oForm.Items.Item("OrdToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				CardTeam = oForm.Items.Item("CardTeam").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				YearPdYN = oForm.Items.Item("YearPdYN").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "       EXEC PS_SD956_03 '";
				sQry += OrdFrDt + "','";
				sQry += OrdToDt + "','";
				sQry += CardType + "','";
				sQry += CardTeam + "','";
				sQry += ItemCode + "','";
				sQry += ItemType + "','";
				sQry += YearPdYN + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_SD956C.ExecuteQuery(sQry);

				oGrid03.Columns.Item(7).RightJustified = true;
				oGrid03.Columns.Item(10).RightJustified = true;
				oGrid03.Columns.Item(13).RightJustified = true;
				oGrid03.Columns.Item(14).RightJustified = true;
				oGrid03.Columns.Item(15).RightJustified = true;
				oGrid03.Columns.Item(18).RightJustified = true;
				oGrid03.Columns.Item(21).RightJustified = true;

				if (oGrid03.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD956_MTX04  회계정보
		/// </summary>
		private void PS_SD956_MTX04()
		{
			int ErrNum = 0;
			string sQry;

			string OrdFrDt;         //수주기간(Fr)
			string OrdToDt;         //수주기간(To)
			string CardType;        //거래처구분
			string CardTeam;        //팀(거래처)
			string ItemCode;        //작번
			string ItemType;        //품목구분(장비/공구)
			string YearPdYN;        //순환품(연간품)

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdFrDt = oForm.Items.Item("OrdFrDt").Specific.Value.ToString().Trim();
				OrdToDt = oForm.Items.Item("OrdToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				CardTeam = oForm.Items.Item("CardTeam").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				YearPdYN = oForm.Items.Item("YearPdYN").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "       EXEC PS_SD956_04 '";
				sQry += OrdFrDt + "','";
				sQry += OrdToDt + "','";
				sQry += CardType + "','";
				sQry += CardTeam + "','";
				sQry += ItemCode + "','";
				sQry += ItemType + "','";
				sQry += YearPdYN + "'";

				oGrid04.DataTable.Clear();
				oDS_PS_SD956D.ExecuteQuery(sQry);

				oGrid04.Columns.Item(4).RightJustified = true;
				oGrid04.Columns.Item(5).RightJustified = true;
				oGrid04.Columns.Item(6).RightJustified = true;
				oGrid04.Columns.Item(7).RightJustified = true;
				oGrid04.Columns.Item(8).RightJustified = true;
				oGrid04.Columns.Item(9).RightJustified = true;
				oGrid04.Columns.Item(10).RightJustified = true;
				oGrid04.Columns.Item(11).RightJustified = true;
				oGrid04.Columns.Item(12).RightJustified = true;
				oGrid04.Columns.Item(13).RightJustified = true;
				oGrid04.Columns.Item(14).RightJustified = true;
				oGrid04.Columns.Item(15).RightJustified = true;
				oGrid04.Columns.Item(16).RightJustified = true;
				oGrid04.Columns.Item(17).RightJustified = true;

				if (oGrid04.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oGrid04.AutoResizeColumns();
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
		/// PS_SD956_FormResize
		/// </summary>
		private void PS_SD956_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 40;
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

				if (oGrid04.Columns.Count > 0)
				{
					oGrid04.AutoResizeColumns();
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
		/// PS_SD956_GetProductionDetail  생산세부정보 조회
		/// </summary>
		private void PS_SD956_GetProductionDetail()
		{
			short loopCount;
			string ItemCode = string.Empty;

			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PS_SD956_01 oTempClass = new PS_SD956_01();

			try
			{
				ProgressBar01.Text = "조회중...";

				for (loopCount = 0; loopCount <= oGrid02.Rows.Count - 1; loopCount++)
				{
					if (oGrid02.Rows.IsSelected(loopCount) == true)
					{
						ItemCode = oGrid02.DataTable.GetValue(1, loopCount).ToString().Trim();
					}
				}

				oTempClass.LoadForm(ItemCode);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oTempClass.oForm.Select();
			}
		}

		/// <summary>
		/// PS_SD956_GetPurchaseDetail  구매세부정보 조회
		/// </summary>
		private void PS_SD956_GetPurchaseDetail()
		{
			short loopCount;
			string ItemCode = string.Empty;

			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PS_SD956_02 oTempClass = new PS_SD956_02();

			try
			{
				ProgressBar01.Text = "조회중...";

				for (loopCount = 0; loopCount <= oGrid03.Rows.Count - 1; loopCount++)
				{
					if (oGrid03.Rows.IsSelected(loopCount) == true)
					{
						ItemCode = oGrid03.DataTable.GetValue(1, loopCount).ToString().Trim();
					}
				}
				
				oTempClass.LoadForm(ItemCode);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oTempClass.oForm.Select();
			}
		}

		/// <summary>
		/// PS_SD956_GetAccountingDetail  회계세부정보 조회
		/// </summary>
		private void PS_SD956_GetAccountingDetail()
		{
			short loopCount;
			string ItemCode = string.Empty;

			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PS_SD956_03 oTempClass = new PS_SD956_03();

			try
			{
				ProgressBar01.Text = "조회중...";

				for (loopCount = 0; loopCount <= oGrid04.Rows.Count - 1; loopCount++)
				{
					if (oGrid04.Rows.IsSelected(loopCount) == true)
					{

						ItemCode = oGrid04.DataTable.GetValue(1, loopCount).ToString().Trim();
					}
				}
				
				oTempClass.LoadForm(ItemCode);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oTempClass.oForm.Select();
			}
		}
	}
}
