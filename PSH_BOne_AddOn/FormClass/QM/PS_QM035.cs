using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// M/G원소재수입검자자료등록
	/// </summary>
	internal class PS_QM035 : PSH_BaseClass
{
		private string oFormUniqueID;
		private SAPbouiCOM.DBDataSource oDS_PS_QM035H; //등록헤더

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM035.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM035_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM035");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				PS_QM035_CreateItems();
				PS_QM035_FormClear();
				PS_QM035_FormItemEnabled();

				oForm.EnableMenu("1283", true);  // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", true);	 // 복제
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
		/// PS_QM035_CreateItems
		/// </summary>
		private void PS_QM035_CreateItems()
		{
			try
			{
				oDS_PS_QM035H = oForm.DataSources.DBDataSources.Item("@PS_QM035H");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM035_FormClear
		/// </summary>
		private void PS_QM035_FormClear()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM035'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oDS_PS_QM035H.SetValue("Code", 0, "1");
				}
				else
				{
					oDS_PS_QM035H.SetValue("Code", 0, DocNum);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM035_FormItemEnabled
		/// </summary>
		private void PS_QM035_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = true;
					oForm.Items.Item("Tk").Enabled = false;
					oForm.Items.Item("Rg").Enabled = false;
					oForm.Items.Item("Br").Enabled = false;
					oForm.Items.Item("Lm").Enabled = false;
					oForm.Items.Item("Hd").Enabled = false;
					oForm.Items.Item("Et").Enabled = false;
					oForm.Items.Item("Ts").Enabled = false;
					oForm.Items.Item("El").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = true;
					oForm.Items.Item("Tk").Enabled = true;
					oForm.Items.Item("Rg").Enabled = true;
					oForm.Items.Item("Br").Enabled = true;
					oForm.Items.Item("Lm").Enabled = true;
					oForm.Items.Item("Hd").Enabled = true;
					oForm.Items.Item("Et").Enabled = true;
					oForm.Items.Item("Ts").Enabled = true;
					oForm.Items.Item("El").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = false;
					oForm.Items.Item("Tk").Enabled = true;
					oForm.Items.Item("Rg").Enabled = true;
					oForm.Items.Item("Br").Enabled = true;
					oForm.Items.Item("Lm").Enabled = true;
					oForm.Items.Item("Hd").Enabled = true;
					oForm.Items.Item("Et").Enabled = true;
					oForm.Items.Item("Ts").Enabled = true;
					oForm.Items.Item("El").Enabled = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM035_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM035_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "LotNo":
						//Lot내역SET
						sQry = "select MAX(a.ItemCode),MAX(d.ItemName),MAX(a.InDate),SUM(B.Quantity),MAX(d.U_Spec1),MAX(U_Spec2) ";
						sQry += "from OBTN a INNER JOIN ITL1 b ON a.ItemCode = b.ItemCode and a.SysNumber = b.SysNumber ";
						sQry += "INNER JOIN OITL c ON c.LogEntry = b.LogEntry ";
						sQry += "INNER JOIN OITM d ON d.ItemCode = a.ItemCode ";
						sQry += "where c.DocType = '59' ";
						sQry += "and left(a.DistNumber,8) Like '" + oDS_PS_QM035H.GetValue("U_LotNo", 0).ToString().Trim() + "' + '%' ";
						oRecordSet.DoQuery(sQry);

						oDS_PS_QM035H.SetValue("U_ItemCode", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_InDate", 0, Convert.ToDateTime(oRecordSet.Fields.Item(2).Value.ToString().Trim()).ToString("yyyyMMdd"));
						oDS_PS_QM035H.SetValue("U_InWgt", 0, oRecordSet.Fields.Item(3).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_Tk", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_Rg", 0, oRecordSet.Fields.Item(5).Value.ToString().Trim());

						//검사기준 SET
						sQry = "SELECT TOP 1 U_S_Tk_P, U_S_Tk_M, U_S_Rg_P, U_S_Rg_M, U_S_Br, U_S_Lm, U_S_Hd, U_S_Et, U_S_Ts_S, U_S_Ts_E, U_S_El ";
						sQry += "From [@PS_QM034H] ";
						sQry += "ORDER BY CODE DESC ";
						oRecordSet.DoQuery(sQry);

						oDS_PS_QM035H.SetValue("U_S_Tk_P", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Tk_M", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Rg_P", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Rg_M", 0, oRecordSet.Fields.Item(3).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Br", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Lm", 0, oRecordSet.Fields.Item(5).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Hd", 0, oRecordSet.Fields.Item(6).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Et", 0, oRecordSet.Fields.Item(7).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Ts_S", 0, oRecordSet.Fields.Item(8).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_Ts_E", 0, oRecordSet.Fields.Item(9).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_S_El", 0, oRecordSet.Fields.Item(10).Value.ToString().Trim());
						oDS_PS_QM035H.SetValue("U_Et", 0, "이상없음"); //기본SET
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
		/// 측청치의 유효값 검사
		/// </summary>
		/// <returns></returns>
		private bool PS_QM035_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			double SPEC;
			double VALUE;
			double VALUE_MIN;
			double VALUE_MAX;
			double Tk;
			double Rg;
			string errMessage = string.Empty;

			try
			{
				Tk = Convert.ToDouble(oForm.Items.Item("Tk").Specific.Value.ToString().Trim());
                Rg = Convert.ToDouble(oForm.Items.Item("Rg").Specific.Value.ToString().Trim());
                //두께
                if (Convert.ToDouble(oDS_PS_QM035H.GetValue("Tk", 0).ToString().Trim()) != 0)
				{
					VALUE_MIN = Tk - Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Tk_M", 0).ToString().Trim());
					VALUE_MAX = Tk + Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Tk_P", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Tk1").Specific.Value.ToString().Trim());
					if (VALUE_MIN > SPEC || VALUE_MAX < SPEC)
					{
                        oForm.Items.Item("Tk1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        errMessage = "두께를 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//폭
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_Rg", 0).ToString().Trim()) != 0)
				{
					VALUE_MIN = Rg - Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Rg_M", 0).ToString().Trim());
					VALUE_MAX = Rg + Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Rg_P", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Rg1").Specific.Value.ToString().Trim());
					if (VALUE_MIN > SPEC || VALUE_MAX < SPEC)
					{
						oForm.Items.Item("Rg").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "폭을 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//Burr
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_Br", 0).ToString().Trim()) != 0)
				{
					VALUE = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Br", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Br").Specific.Value.ToString().Trim());
					if (VALUE < SPEC)
					{
						oForm.Items.Item("Br").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "Burr를 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//조도
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_Lm", 0).ToString().Trim()) != 0)
				{
					VALUE = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Lm", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Lm").Specific.Value.ToString().Trim());
					if (VALUE < SPEC)
					{
						oForm.Items.Item("Lm").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "조도를 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//경도
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_Hd", 0).ToString().Trim()) != 0)
				{
					VALUE = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Hd", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Hd").Specific.Value.ToString().Trim());
					if (VALUE < SPEC)
					{
						oForm.Items.Item("Hd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "경도를 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//인장강도
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_Ts", 0).ToString().Trim()) != 0)
				{
					VALUE_MIN = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Ts_S", 0).ToString().Trim());
					VALUE_MAX = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_Ts_E", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("Ts").Specific.Value.ToString().Trim());
					if (VALUE_MIN > SPEC || VALUE_MAX < SPEC)
					{
						oForm.Items.Item("Ts").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "인장강도를 확인하여 주십시오.";
						throw new Exception();
					}
				}
				//연신율
				if (Convert.ToDouble(oDS_PS_QM035H.GetValue("U_El", 0).ToString().Trim()) != 0)
				{
					VALUE = Convert.ToDouble(oDS_PS_QM035H.GetValue("U_S_El", 0).ToString().Trim());
					SPEC = Convert.ToDouble(oForm.Items.Item("El").Specific.Value.ToString().Trim());
					if (VALUE > SPEC)
					{
						oForm.Items.Item("El").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "연신율을 확인하여 주십시오.";
						throw new Exception();
					}
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
				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
							if (PS_QM035_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
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
							PS_QM035_FormItemEnabled();
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "LotNo")
						{
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								if (!string.IsNullOrEmpty(oForm.Items.Item("LotNo").Specific.Value.ToString().Trim()))
								{
									sQry = "select Count(*) From [@PS_QM035H] Where U_LotNo = '" + oDS_PS_QM035H.GetValue("U_LotNo", 0).ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
									{
										PS_QM035_FlushToItemValue(pVal.ItemUID, 0, "");
									}
									else
									{
										errMessage = "이미 등록된 Lot입니다.";
										throw new Exception();
									}
                                }
								else
								{
									oForm.Items.Item("ItemCode").Specific.Value = "";
									oForm.Items.Item("ItemName").Specific.Value = "";
									oForm.Items.Item("InDate").Specific.Value = "";
									oForm.Items.Item("InWgt").Specific.Value = "0";
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
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
				BubbleEvent = false;
                return;
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM035H);
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
						case "1283": //제거
							break;
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
							PS_QM035_FormItemEnabled();
                            oForm.Items.Item("Tk").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
						case "1282": //추가
							PS_QM035_FormItemEnabled();
							PS_QM035_FormClear();
							oForm.Items.Item("C_Tk").Specific.Value = "";
							oForm.Items.Item("C_Rg").Specific.Value = "";
                            oForm.Items.Item("LotNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM035_FormItemEnabled();
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
