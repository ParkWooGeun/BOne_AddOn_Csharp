using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 방산부품 원재료 입고 등록
	/// </summary>
	internal class PS_QM082 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.DBDataSource oDS_PS_QM082H; //등록헤더

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM082.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM082_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM082");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM082_CreateItems();
				PS_QM082_ComboBox_Setting();
				PS_QM082_FormClear();

				oForm.EnableMenu("1283", true);  // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_QM082_CreateItems
		/// </summary>
		private void PS_QM082_CreateItems()
		{
			try
			{
				oDS_PS_QM082H = oForm.DataSources.DBDataSources.Item("@PS_QM082H");
				oDS_PS_QM082H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM082_ComboBox_Setting
		/// </summary>
		private void PS_QM082_ComboBox_Setting()
		{
			try
			{
				// 납품완료여부(Y/N)
				oForm.Items.Item("FinishYN").Specific.ValidValues.Add("N", "납품미완료(N)");
				oForm.Items.Item("FinishYN").Specific.ValidValues.Add("Y", "납품완료(Y)");
				oForm.Items.Item("FinishYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

				// 검사완료여부(Y/N)
				oForm.Items.Item("TestYN").Specific.ValidValues.Add("N", "검사미완료(N)");
				oForm.Items.Item("TestYN").Specific.ValidValues.Add("Y", "검사완료(Y)");
				oForm.Items.Item("TestYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM082_FormClear
		/// </summary>
		private void PS_QM082_FormClear()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM082'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oDS_PS_QM082H.SetValue("Code", 0, "1");
				}
				else
				{
					oDS_PS_QM082H.SetValue("Code", 0, DocNum);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM082_FormItemEnabled
		/// </summary>
		private void PS_QM082_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("BaseCode").Enabled = true;
					oForm.Items.Item("LotNo").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BaseCode").Enabled = true;
					oForm.Items.Item("LotNo").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("BaseCode").Enabled = false;
					oForm.Items.Item("LotNo").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM082_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_QM082_FlushToItemValue(string oUID)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "BaseCode":
						//검사사양서내역
						sQry = "select U_ItemCode, U_ItemName, U_CItemCod, U_CItemNam, U_StdNum,";
						sQry += " U_C_min, U_C_max, U_Mn_min, U_Mn_max, U_P_min, U_S_min, U_Si_min, U_Si_max, U_Ni_min, U_Ni_max,";
						sQry += " U_Cr_min, U_Cr_max, U_Mo_min, U_Mo_max, U_Fe_min, U_Cu_min, U_Cu_max, U_Mg_min, U_Mg_max, U_Zn_min,";
						sQry += " U_Ti_min, U_ZrTi_min, U_Other, U_Al, U_TS_min,U_TS_max, U_YS_min, U_YS_max, U_EL_min, U_EL_max,";
						sQry += " U_HARD_min, U_HARD_max, U_GS_min, U_GS_max, U_Clean, U_METCH,U_EACH, U_TOTAL";
						sQry += " From [@PS_QM081H] ";
						sQry += " where Code = '" + oDS_PS_QM082H.GetValue("U_BaseCode", 0).ToString().Trim() + "' ";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("ItemCode").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
						oForm.Items.Item("CItemCod").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
						oForm.Items.Item("CItemNam").Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
						oForm.Items.Item("StdNum").Specific.Value = oRecordSet.Fields.Item(4).Value.ToString().Trim();
						oForm.Items.Item("C_min").Specific.Value = oRecordSet.Fields.Item(5).Value.ToString().Trim();
						oForm.Items.Item("C_max").Specific.Value = oRecordSet.Fields.Item(6).Value.ToString().Trim();
						oForm.Items.Item("Mn_min").Specific.Value = oRecordSet.Fields.Item(7).Value.ToString().Trim();
						oForm.Items.Item("Mn_max").Specific.Value = oRecordSet.Fields.Item(8).Value.ToString().Trim();
						oForm.Items.Item("P_min").Specific.Value = oRecordSet.Fields.Item(9).Value.ToString().Trim();
						oForm.Items.Item("S_min").Specific.Value = oRecordSet.Fields.Item(10).Value.ToString().Trim();
						oForm.Items.Item("Si_min").Specific.Value = oRecordSet.Fields.Item(11).Value.ToString().Trim();
						oForm.Items.Item("Si_max").Specific.Value = oRecordSet.Fields.Item(12).Value.ToString().Trim();
						oForm.Items.Item("Ni_min").Specific.Value = oRecordSet.Fields.Item(13).Value.ToString().Trim();
						oForm.Items.Item("Ni_max").Specific.Value = oRecordSet.Fields.Item(14).Value.ToString().Trim();
						oForm.Items.Item("Cr_min").Specific.Value = oRecordSet.Fields.Item(15).Value.ToString().Trim();
						oForm.Items.Item("Cr_max").Specific.Value = oRecordSet.Fields.Item(16).Value.ToString().Trim();
						oForm.Items.Item("Mo_min").Specific.Value = oRecordSet.Fields.Item(17).Value.ToString().Trim();
						oForm.Items.Item("Mo_max").Specific.Value = oRecordSet.Fields.Item(18).Value.ToString().Trim();
						oForm.Items.Item("Fe_min").Specific.Value = oRecordSet.Fields.Item(19).Value.ToString().Trim();
						oForm.Items.Item("Cu_min").Specific.Value = oRecordSet.Fields.Item(20).Value.ToString().Trim();
						oForm.Items.Item("Cu_max").Specific.Value = oRecordSet.Fields.Item(21).Value.ToString().Trim();
						oForm.Items.Item("Mg_min").Specific.Value = oRecordSet.Fields.Item(22).Value.ToString().Trim();
						oForm.Items.Item("Mg_max").Specific.Value = oRecordSet.Fields.Item(23).Value.ToString().Trim();
						oForm.Items.Item("Zn_min").Specific.Value = oRecordSet.Fields.Item(24).Value.ToString().Trim();
						oForm.Items.Item("Ti_min").Specific.Value = oRecordSet.Fields.Item(25).Value.ToString().Trim();
						oForm.Items.Item("ZrTi_min").Specific.Value = oRecordSet.Fields.Item(26).Value.ToString().Trim();
						oForm.Items.Item("Other").Specific.Value = oRecordSet.Fields.Item(27).Value.ToString().Trim();
						oForm.Items.Item("Al").Specific.Value = oRecordSet.Fields.Item(28).Value.ToString().Trim();
						oForm.Items.Item("EACH").Specific.Value = oRecordSet.Fields.Item(41).Value.ToString().Trim();
						oForm.Items.Item("TOTAL").Specific.Value = oRecordSet.Fields.Item(42).Value.ToString().Trim();

						oForm.Items.Item("TS_min").Specific.Value = oRecordSet.Fields.Item(29).Value.ToString().Trim();
						oForm.Items.Item("TS_max").Specific.Value = oRecordSet.Fields.Item(30).Value.ToString().Trim();
						oForm.Items.Item("YS_min").Specific.Value = oRecordSet.Fields.Item(31).Value.ToString().Trim();
						oForm.Items.Item("YS_max").Specific.Value = oRecordSet.Fields.Item(32).Value.ToString().Trim();
						oForm.Items.Item("EL_min").Specific.Value = oRecordSet.Fields.Item(33).Value.ToString().Trim();
						oForm.Items.Item("EL_max").Specific.Value = oRecordSet.Fields.Item(34).Value.ToString().Trim();
						oForm.Items.Item("HARD_min").Specific.Value = oRecordSet.Fields.Item(35).Value.ToString().Trim();
						oForm.Items.Item("HARD_max").Specific.Value = oRecordSet.Fields.Item(36).Value.ToString().Trim();
						oForm.Items.Item("GS_min").Specific.Value = oRecordSet.Fields.Item(37).Value.ToString().Trim();
						oForm.Items.Item("GS_max").Specific.Value = oRecordSet.Fields.Item(38).Value.ToString().Trim();
						oForm.Items.Item("Clean").Specific.Value = oRecordSet.Fields.Item(39).Value.ToString().Trim();
						oForm.Items.Item("METCH").Specific.Value = oRecordSet.Fields.Item(40).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_QM082_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM082_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_QM082H.GetValue("U_LotNo", 0).ToString().Trim()))
				{
					errMessage = "otNo는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM082H.GetValue("U_BaseCode", 0).ToString().Trim()))
				{
					errMessage = "원재료선택은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM082H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					errMessage = "입고일자는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM082H.GetValue("U_Qty", 0).ToString().Trim()))
				{
					errMessage = "수량은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM082H.GetValue("U_Weight", 0).ToString().Trim()))
				{
					errMessage = "중량은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}

				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM082_Check_Exist
		/// </summary>
		/// <returns></returns>
		private bool PS_QM082_Check_Exist()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "select * from [@PS_QM082H] where U_DocDate = '" + oDS_PS_QM082H.GetValue("U_DocDate", 0).ToString().Trim() + "' and ";
				sQry += " U_LotNo = '" + oDS_PS_QM082H.GetValue("U_LotNo", 0).ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount != 0)
				{
					errMessage = "이미 동일한 입고일자에 LotNo가 존재합니다. 확인하여 주십시오.";
					throw new Exception();
				}

				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
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
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
				//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
				//    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
				//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
				//    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
				//    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
				//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_Drag: //39
				//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
							if (PS_QM082_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								if (PS_QM082_Check_Exist() == false)
								{
									BubbleEvent = false;
									return;
								}
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							PS_QM082_FormItemEnabled();
							PS_QM082_FlushToItemValue("BaseCode");
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "BaseCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("BaseCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "BaseCode")
						{
							PS_QM082_FlushToItemValue(pVal.ItemUID);
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM082H);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				oForm.Freeze(true);
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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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
							PS_QM082_FormItemEnabled();
							oForm.Items.Item("LotNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM082_FormItemEnabled();
							PS_QM082_FormClear();
							oDS_PS_QM082H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							oForm.Items.Item("LotNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM082_FormItemEnabled();
							PS_QM082_FlushToItemValue("BaseCode");
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// FormDataEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="BusinessObjectInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
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
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
