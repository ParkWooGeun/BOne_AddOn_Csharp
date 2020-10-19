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
		public string oFormUniqueID01;
		public SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_CO170H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO170L; //등록라인

	//	private string oDocType01;
	//	private string oDoType01;
		public string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		public string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		public int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		//public override void LoadForm(string oFromDocEntry01, string oFromDocType01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO170.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_CO170_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO170");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";                // UDO방식일때

				oForm.Freeze(true);
			//	oDocType01 = oFromDocType01;
			//	oDoType01 = oFromDocType01;
				PS_CO170_CreateItems();
				PS_CO170_ComboBox_Setting();
				PS_CO170_EnableMenus();
				PS_CO170_AddMatrixRow(0, true);             //UDO방식
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
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				switch (pval.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:               //1
						Raise_EVENT_ITEM_PRESSED(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                   //2
						Raise_EVENT_KEY_DOWN(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:               //5
						Raise_EVENT_COMBO_SELECT(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:                      //6
						Raise_EVENT_CLICK(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:               //7
						Raise_EVENT_DOUBLE_CLICK(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:        //8
						Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:                   //10
						Raise_EVENT_VALIDATE(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                //11
						Raise_EVENT_MATRIX_LOAD(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:              //18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:            //19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                //20
						Raise_EVENT_RESIZE(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:           //27
						Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                  //3
						Raise_EVENT_GOT_FOCUS(FormUID, ref pval, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                 //4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                //17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pval, ref BubbleEvent);
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
		/// Raise_FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.MenuUID)
					{
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pval, ref BubbleEvent);
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.MenuUID)
					{
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pval, ref BubbleEvent);
							break;
						case "1281":                            //찾기
							PS_CO170_FormItemEnabled();         //UDO방식
							oForm.Items.Item("U_ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282":                            //추가
							PS_CO170_FormItemEnabled();         //UDO방식
							PS_CO170_AddMatrixRow(0, true);     //UDO방식
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
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
		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE))
							{
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "Mat01")
					{
						if (pval.Row > 0)
						{
							oLastItemUID01 = pval.ItemUID;
							oLastColUID01 = pval.ColUID;
							oLastColRow01 = pval.Row;
						}
					}
					else
					{
						oLastItemUID01 = pval.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
					}
				}
				else if (pval.BeforeAction == false)
				{
					if (pval.ItemUID == "Mat01")
					{
						if (pval.Row > 0)
						{
							oLastItemUID01 = pval.ItemUID;
							oLastColUID01 = pval.ColUID;
							oLastColRow01 = pval.Row;
						}
					}
					else
					{
						oLastItemUID01 = pval.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			string ItmMsort = string.Empty;
			string ItmBsort = string.Empty;
			string Code = string.Empty;

			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "1")
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
				else if (pval.BeforeAction == false)
				{
					if (pval.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pval.ActionSuccess == true)
							{
								PS_CO170_FormItemEnabled();
								PS_CO170_AddMatrixRow(oMat01.RowCount, true);   //UDO방식일때
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (pval.ActionSuccess == true)
							{
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pval.ActionSuccess == true)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.CharPressed == 9)
					{
						if (pval.ItemUID == "U_ItmBsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("U_ItmBsort").Specific.VALUE))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
								BubbleEvent = false;
							}
						}
						if (pval.ItemUID == "ItmMsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.VALUE))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
								BubbleEvent = false;
							}
						}
						if (pval.ItemUID == "Mat01")
						{
							if (pval.ColUID == "PO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
									BubbleEvent = false;
								}
							}
							if (pval.ColUID == "MPO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
									BubbleEvent = false;
								}
							}
						}
					}
				}
				else if (pval.BeforeAction == false)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
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
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "Mat01")
					{
						if (pval.Row > 0)
						{
							oMat01.SelectRow(pval.Row, true, false);
						}
					}
				}
				else if (pval.BeforeAction == false)
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
		/// 
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
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
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			int i = 0;
			int ErrNum = 0;
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pval.BeforeAction == true)
				{
					if (pval.ItemChanged == true)
					{
						if ((pval.ItemUID == "Mat01"))
						{
							if (pval.ColUID == "PO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE.ToString().Trim()))
								{
									ErrNum = 1;
									throw new Exception();
								}
								for (i = 1; i <= oMat01.RowCount; i++)
								{
									// 현재 선택되어있는 행이 아니면
									if (pval.Row != i)
									{
										if ((oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE == oMat01.Columns.Item("PO").Cells.Item(i).Specific.VALUE))
										{
											PSH_Globals.SBO_Application.MessageBox("동일한 항목이 존재합니다.");
											oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE = "";
											ErrNum = 0;
											throw new Exception();
										}
									}
								}

								sQry = "EXEC PS_CO170_01 '" + oMat01.Columns.Item("PO").Cells.Item(pval.Row).Specific.VALUE.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
								{
									oDS_PS_CO170L.SetValue("U_PO", pval.Row - 1, oRecordSet.Fields.Item("PO").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_POEntry", pval.Row - 1, oRecordSet.Fields.Item("POEntry").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_POLine", pval.Row - 1, oRecordSet.Fields.Item("POLine").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_Sequence", pval.Row - 1, oRecordSet.Fields.Item("Sequence").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_ItemCode", pval.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_ItemName", pval.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_CpCode", pval.Row - 1, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_CpName", pval.Row - 1, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
									oRecordSet.MoveNext();
								}
								if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(oDS_PS_CO170L.GetValue("U_PO", pval.Row - 1).ToString().Trim()))
								{
									PS_CO170_AddMatrixRow((pval.Row));
								}
								oMat01.LoadFromDataSource();
								oMat01.AutoResizeColumns();
							}
							else if (pval.ColUID == "MPO")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE.ToString().Trim()))
								{
									ErrNum = 2;
									throw new Exception();
								}
								sQry = "EXEC PS_CO170_01 '" + oMat01.Columns.Item("MPO").Cells.Item(pval.Row).Specific.VALUE.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
								{
									oDS_PS_CO170L.SetValue("U_MPO", pval.Row - 1, oRecordSet.Fields.Item("PO").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MPOEntry", pval.Row - 1, oRecordSet.Fields.Item("POEntry").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MPOLine", pval.Row - 1, oRecordSet.Fields.Item("POLine").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MSequenc", pval.Row - 1, oRecordSet.Fields.Item("Sequence").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MItemCod", pval.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MItemNam", pval.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MCpCode", pval.Row - 1, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
									oDS_PS_CO170L.SetValue("U_MCpName", pval.Row - 1, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
									oRecordSet.MoveNext();
								}
								oMat01.LoadFromDataSource();
								oMat01.AutoResizeColumns();
							}
							else
							{
								oDS_PS_CO170L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.VALUE.ToString().Trim());
							}

							oMat01.LoadFromDataSource();
							oMat01.AutoResizeColumns();
							oForm.Update();
							oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							if ((pval.ItemUID == "U_ItmBsort"))
							{
								oDS_PS_CO170H.SetValue(pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
								oDS_PS_CO170H.SetValue("U_ItmBName", 0, dataHelpClass.GetValue("SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" + oForm.Items.Item("U_ItmBsort").Specific.VALUE.ToString().Trim() + "'", 0, 1));
							}
							else if ((pval.ItemUID == "ItmMsort"))
							{
								oDS_PS_CO170H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE);
								oDS_PS_CO170H.SetValue("U_ItmMName", 0, dataHelpClass.GetValue("SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("U_ItmBsort").Specific.VALUE.ToString().Trim() + "' And U_Code = '" + oForm.Items.Item(pval.ItemUID).Specific.VALUE.ToString().Trim() + "'", 0, 1));
							}
							else
							{
								oDS_PS_CO170H.SetValue("U_" + pval.ItemUID, 0, oForm.Items.Item(pval.ItemUID).Specific.VALUE.ToString().Trim());
							}
						}
					}
				}
				else if (pval.BeforeAction == false)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
				{
					PS_CO170_FormItemEnabled();
					PS_CO170_AddMatrixRow(oMat01.VisualRowCount);                   //UDO방식
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
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
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
		/// Raise_EVENT_CHOOSE_FROM_LIST
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
				{
					if ((pval.ItemUID == "U_ItmBsort" | pval.ItemUID == "ItmBName"))
					{
						dataHelpClass.PSH_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO170H", "U_ItmBsort,U_ItmBName", "", 0, "", "", "");
					}
					if ((pval.ItemUID == "ItmMsort" | pval.ItemUID == "ItmMName"))
					{
						dataHelpClass.PSH_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO170H", "U_ItmMsort,U_ItmMName", "", 0, "", "", "");
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
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.ItemUID == "Mat01")
				{
					if (pval.Row > 0)
					{
						oLastItemUID01 = pval.ItemUID;
						oLastColUID01 = pval.ColUID;
						oLastColRow01 = pval.Row;
					}
				}
				else
				{
					oLastItemUID01 = pval.ItemUID;
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO170H); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO170L); //메모리 해제
					SubMain.Remove_Forms(oFormUniqueID01);
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			int i = 0;
			try
			{
				if ((oLastColRow01 > 0))
				{
					if (pval.BeforeAction == true)
					{
					}
					else if (pval.BeforeAction == false)
					{
						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
						}
						oMat01.FlushToDataSource();
						oDS_PS_CO170L.RemoveRecord(oDS_PS_CO170L.Size - 1);
						oMat01.LoadFromDataSource();
						if (oMat01.RowCount == 0)
						{
							PS_CO170_AddMatrixRow(0);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_CO170L.GetValue("U_PO", oMat01.RowCount - 1).ToString().Trim()))
							{
								PS_CO170_AddMatrixRow(oMat01.RowCount);
							}
						}
						oForm.Update();
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
		/// PS_CO170_CreateItems
		/// </summary>
		/// <returns></returns>
		private bool PS_CO170_CreateItems()
		{
			bool functionReturnValue = false;
			try
			{
				oForm.Freeze(true);

				oDS_PS_CO170H = oForm.DataSources.DBDataSources.Item("@PS_CO170H");
				oDS_PS_CO170L = oForm.DataSources.DBDataSources.Item("@PS_CO170L");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				functionReturnValue = false;
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}

			oForm.Freeze(false);
			return functionReturnValue;
		}

		/// <summary>
		/// PS_CO170_ComboBox_Setting
		/// </summary>
		public void PS_CO170_ComboBox_Setting()
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
		public void PS_CO170_FormItemEnabled()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
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
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
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
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
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

		public void PS_CO170_AddMatrixRow(int oRow, bool RowIserted = false)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)   //행추가여부
				{
					oDS_PS_CO170L.InsertRecord((oRow));
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
		/// 
		/// </summary>
		/// <returns></returns>
		public bool PS_CO170_DataValidCheck()
		{
			bool functionReturnValue = false;
			int i = 0;
			int ErrNum = 0;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("U_ItmBsort").Specific.VALUE.ToString().Trim()))
				{
					oForm.Items.Item("U_ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.VALUE.ToString().Trim()))
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
					if ((string.IsNullOrEmpty(oMat01.Columns.Item("PO").Cells.Item(i).Specific.VALUE.ToString().Trim())))
					{
						oMat01.Columns.Item("PO").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						ErrNum = 4;
						throw new Exception();
					}
					if ((Convert.ToDouble(oMat01.Columns.Item("MPO").Cells.Item(i).Specific.VALUE.ToString().Trim()) <= 0))
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
				functionReturnValue = false;
			}
			return functionReturnValue;
		}
	}
}
