using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 재공완료품대체관리등록
	/// </summary>
	internal class PS_CO170 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_CO170H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO170L; //등록라인
		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO170.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_CO170_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO170"); // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				PS_CO170_CreateItems();
				PS_CO170_ComboBox_Setting();
				PS_CO170_EnableMenus();
				PS_CO170_AddMatrixRow(0, true);
				PSH_Globals.ExecuteEventFilter(typeof(PS_CO170));
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
		/// PS_CO170_CreateItems
		/// </summary>
		/// <returns></returns>
		private void PS_CO170_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_CO170H = oForm.DataSources.DBDataSources.Item("@PS_CO170H");
				oDS_PS_CO170L = oForm.DataSources.DBDataSources.Item("@PS_CO170L");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();
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
		/// PS_CO170_ComboBox_Setting
		/// </summary>
		private void PS_CO170_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//콤보에 기본값설정
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

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
		/// PS_CO170_FormItemEnabled
		/// </summary>
		private void PS_CO170_FormItemEnabled()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					// 각모드에따른 아이템설정
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("U_ItmBsort").Enabled = true;
					oForm.Items.Item("ItmMsort").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
					oMat01.AutoResizeColumns();
					oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
					oForm.EnableMenu("1281", true);             //찾기
					oForm.EnableMenu("1282", false);            //추가

				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("U_ItmBsort").Enabled = true;
					oForm.Items.Item("ItmMsort").Enabled = true;
					oForm.Items.Item("Comment").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oMat01.AutoResizeColumns();
					oForm.EnableMenu("1281", false);
					oForm.EnableMenu("1282", true);
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true);             //찾기
					oForm.EnableMenu("1282", true);             //추가
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("U_ItmBsort").Enabled = false;
					oForm.Items.Item("ItmMsort").Enabled = false;
					oMat01.AutoResizeColumns();
					oForm.EnableMenu("1281", true);
					oForm.EnableMenu("1282", false);
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
		/// 행추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_CO170_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)   //행추가여부
				{
					oDS_PS_CO170L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_CO170L.Offset = oRow;
				oDS_PS_CO170L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat01.LoadFromDataSource();
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
		/// PS_CO170_EnableMenus
		/// </summary>
		private void PS_CO170_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, true, false);   // 메뉴설정
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
		/// DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_CO170_DataValidCheck()
		{
			bool functionReturnValue = false;
			int i;
			int ErrNum = 0;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("U_ItmBsort").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("U_ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("ItmMsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					ErrNum = 2;
					throw new Exception();
				}

				if (oMat01.VisualRowCount == 1)
				{
					ErrNum = 3;
					throw new Exception();
				}
				for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat01.Columns.Item("PO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						ErrNum = 4;
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat01.Columns.Item("MPO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						ErrNum = 5;
						throw new Exception();
					}
				}
				oDS_PS_CO170L.RemoveRecord(oDS_PS_CO170L.Size - 1);
				oMat01.LoadFromDataSource();
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("품목대분류는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("품목중분류는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 4)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("작지문서라인은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 5)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("이동작지문서라인은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}

			return functionReturnValue;
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
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:               //1
						Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                   //2
						Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:               //5
						//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:                      //6
						Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:               //7
						//Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:        //8
						//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:                   //10
						Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                //11
						Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:              //18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:            //19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                //20
						//Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:           //27
						Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                  //3
						Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                 //4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                //17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string ItmMsort;
			string ItmBsort;
			string Code;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_CO170_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							ItmBsort = oDS_PS_CO170H.GetValue("U_ItmBsort", 0).ToString().Trim();
							ItmMsort = oDS_PS_CO170H.GetValue("U_ItmMsort", 0).ToString().Trim();
							Code = ItmBsort + ItmMsort;
							oDS_PS_CO170H.SetValue("Code", 0, Code);
							oDS_PS_CO170H.SetValue("Name", 0, Code);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_CO170_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_CO170_FormItemEnabled();
								PS_CO170_AddMatrixRow(oMat01.RowCount, true);   //UDO방식일때
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_CO170_FormItemEnabled();
							}
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "U_ItmBsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("U_ItmBsort").Specific.Value))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "ItmMsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.Value))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							if (pVal.ColUID == "MPO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
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
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.ItemUID == "Mat01")
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
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;

							oMat01.SelectRow(pVal.Row, true, false);
						}
					}
					else
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
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
			int i;
			int ErrNum = 0;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if ((pVal.ItemUID == "Mat01"))
						{
							if (pVal.ColUID == "PO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									ErrNum = 1;
									throw new Exception();
								}
								for (i = 1; i <= oMat01.RowCount; i++)
								{
									// 현재 선택되어있는 행이 아니면
									if (pVal.Row != i)
									{
										if ((oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("PO").Cells.Item(i).Specific.Value))
										{
											PSH_Globals.SBO_Application.MessageBox("동일한 항목이 존재합니다.");
											oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value = "";
											ErrNum = 0;
											throw new Exception();
										}
									}
								}

								sQry = "EXEC PS_CO170_01 '" + oMat01.Columns.Item("PO").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
								{
									oDS_PS_CO170L.SetValue("U_PO", pVal.Row - 1, oRecordSet.Fields.Item("PO").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_POEntry", pVal.Row - 1, oRecordSet.Fields.Item("POEntry").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_POLine", pVal.Row - 1, oRecordSet.Fields.Item("POLine").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_Sequence", pVal.Row - 1, oRecordSet.Fields.Item("Sequence").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_CpName", pVal.Row - 1, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
									oRecordSet.MoveNext();
								}
								if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_CO170L.GetValue("U_PO", pVal.Row - 1).ToString().Trim()))
								{
									PS_CO170_AddMatrixRow(pVal.Row, false);
								}
								oMat01.LoadFromDataSource();
								oMat01.AutoResizeColumns();
							}
							else if (pVal.ColUID == "MPO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									ErrNum = 2;
									throw new Exception();
								}
								sQry = "EXEC PS_CO170_01 '" + oMat01.Columns.Item("MPO").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
								{
									oDS_PS_CO170L.SetValue("U_MPO", pVal.Row - 1, oRecordSet.Fields.Item("PO").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MPOEntry", pVal.Row - 1, oRecordSet.Fields.Item("POEntry").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MPOLine", pVal.Row - 1, oRecordSet.Fields.Item("POLine").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MSequenc", pVal.Row - 1, oRecordSet.Fields.Item("Sequence").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MItemCod", pVal.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MItemNam", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MCpCode", pVal.Row - 1, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MCpName", pVal.Row - 1, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
									oRecordSet.MoveNext();
								}
								oMat01.LoadFromDataSource();
								oMat01.AutoResizeColumns();
							}
							else
							{
								oDS_PS_CO170L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}

							oMat01.LoadFromDataSource();
							oMat01.AutoResizeColumns();
							oForm.Update();
							oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							if ((pVal.ItemUID == "U_ItmBsort"))
							{
								oDS_PS_CO170H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
								oDS_PS_CO170H.SetValue("U_ItmBName", 0, dataHelpClass.GetValue("SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" + oForm.Items.Item("U_ItmBsort").Specific.Value.ToString().Trim() + "'", 0, 1));
							}
							else if ((pVal.ItemUID == "ItmMsort"))
							{
								oDS_PS_CO170H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
								oDS_PS_CO170H.SetValue("U_ItmMName", 0, dataHelpClass.GetValue("SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("U_ItmBsort").Specific.Value.ToString().Trim() + "' And U_Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", 0, 1));
							}
							else
							{
								oDS_PS_CO170H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
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
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("작지문서라인을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("대체작지문서라인을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Freeze(false);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_CO170_FormItemEnabled();
					PS_CO170_AddMatrixRow(oMat01.VisualRowCount, false);
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
		/// Raise_EVENT_CHOOSE_FROM_LIST
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "U_ItmBsort" || pVal.ItemUID == "ItmBName")
					{
						dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_CO170H", "U_ItmBsort,U_ItmBName", "", 0, "", "", "");
					}
					if (pVal.ItemUID == "ItmMsort" || pVal.ItemUID == "ItmMName")
					{
						dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_CO170H", "U_ItmMsort,U_ItmMName", "", 0, "", "", "");
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO170H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO170L);
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
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i = 0;

			try
			{
				oForm.Freeze(true);

				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{

					}
					else if (pVal.BeforeAction == false)
					{
						if (oMat01.RowCount != oMat01.VisualRowCount)
						{
							oMat01.FlushToDataSource();

							while (i <= oDS_PS_CO170L.Size - 1)
							{
								if (string.IsNullOrEmpty(oDS_PS_CO170L.GetValue("U_PO", i)))
								{
									oDS_PS_CO170L.RemoveRecord(i);
									i = 0;
								}
								else
								{
									i += 1;
								}
							}

							for (i = 0; i <= oDS_PS_CO170L.Size; i++)
							{
								oDS_PS_CO170L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
							}

							oMat01.LoadFromDataSource();
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							PS_CO170_AddMatrixRow(oMat01.RowCount, false);
							break;
						case "1281": //찾기
							PS_CO170_FormItemEnabled();
							oForm.Items.Item("U_ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_CO170_FormItemEnabled();
							PS_CO170_AddMatrixRow(0, true);
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
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				if (BusinessObjectInfo.BeforeAction == true)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}

				if (pVal.ItemUID == "Mat01")
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
	}
}
