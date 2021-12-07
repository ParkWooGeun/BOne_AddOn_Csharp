using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	///  화면 및 권한 변경이력조회
	/// </summary>
	internal class PS_SY020 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Grid oGrid01;
		private string oLastItemUID01;   // 클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;    // 마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;       // 마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SY020.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				// 매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_SY020_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_SY020");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_SY020_CreateItems();
				PS_SY020_ComboBox_Setting();
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
				oForm.ActiveItem = "BPLId";   //최초 커서위치
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// PS_SY020_CreateItems
		/// </summary>
		private void PS_SY020_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.UserDataSources.Add("DocDatefr", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("DocDatefr").Specific.DataBind.SetBound(true, "", "DocDatefr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

				oForm.DataSources.UserDataSources.Add("DocDateto", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("DocDateto").Specific.DataBind.SetBound(true, "", "DocDateto");
				oForm.DataSources.UserDataSources.Item("DocDateto").Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SY020_ComboBox_Setting
		/// </summary>
		private void PS_SY020_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] order by BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				// Addon/Core 구분
				oForm.Items.Item("Type").Specific.ValidValues.Add("A", "에드온화면");
				oForm.Items.Item("Type").Specific.ValidValues.Add("C", "코어화면");
				oForm.Items.Item("Type").DisplayDesc = true;

				// 변경타입
				oForm.Items.Item("MType").Specific.ValidValues.Add("N", "신규");
				oForm.Items.Item("MType").Specific.ValidValues.Add("M", "변경");
				oForm.Items.Item("MType").Specific.ValidValues.Add("C", "부서이동");
				oForm.Items.Item("MType").DisplayDesc = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SY020_MTX01  메트릭스에 데이터 로드
		/// </summary>
		private void PS_SY020_MTX01()
		{
			int ErrNum = 0;
			string Query01;
			string UserID;
			string DocDateFr;
			string DocDateTo;
			string String;
			string type;
			string mType;
			SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				UserID = oForm.Items.Item("UserID").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDatefr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateto").Specific.Value.ToString().Trim();
				String = oForm.Items.Item("String").Specific.Value.ToString().Trim();
				type = oForm.Items.Item("Type").Specific.Value.ToString().Trim();
				mType = oForm.Items.Item("MType").Specific.Value.ToString().Trim();

				Query01 = "EXEC PS_SY020_01 '" + UserID + "','" + DocDateFr + "','" + DocDateTo + "','" + String + "','" + type + "','" + mType + "'";

				oGrid01.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01);
				oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				if (oGrid01.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
				oGrid01.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				if (ProgBar01 != null)
				{
					ProgBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				switch (pVal.EventType)
				{
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: // 1
                        Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:     // 2
                    //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:    // 3
                        Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:   // 4
                        break;
                    //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: // 5
                    //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_CLICK:        // 6
                    //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: // 7
                    //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: // 8
                    //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_VALIDATE: // 10
                    //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: // 11
                    //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: // 17
                        Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: // 18
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: // 19
                        break;
                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:  // 20
                    //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: // 27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
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
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SY020_MTX01();
						}
					}
					if (pVal.ItemUID == "Button02")
					{
						oForm.Close();
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
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
					SubMain.Remove_Forms(oFormUniqueID01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
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
		/// Raise_FormMenuEvent
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
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
		/// 
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:  // 33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:   // 34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:// 35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:// 36
							break;
					}
					////BeforeAction = False
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:  // 33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:   // 34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:// 35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:// 36
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
	}
}
