using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 원소재성분입력
	/// </summary>
	internal class PS_QM030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_QM030H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM030L; //등록라인

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM030_CreateItems();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
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
		/// PS_QM030_CreateItems
		/// </summary>
		private void PS_QM030_CreateItems()
		{
			try
			{
				oDS_PS_QM030H = oForm.DataSources.DBDataSources.Item("@PS_QM030H");
				oDS_PS_QM030L = oForm.DataSources.DBDataSources.Item("@PS_QM030L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM030_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_QM030_FlushToItemValue(string oUID)
		{
			int i;
			int j;
			double Counts;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "LotNo":
						Counts = oForm.Items.Item("ItemCode").Specific.ValidValues.Count;
						if (Counts != 0)
						{
							for (j = 0; j <= Convert.ToDouble(Counts) - 1; j++)
							{
								oForm.Items.Item("ItemCode").Specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						oDS_PS_QM030H.SetValue("U_ItemCode", 0, "");
						oDS_PS_QM030H.SetValue("U_ChemC_Fe", 0, "0");
						oDS_PS_QM030H.SetValue("U_ChemC_P", 0, "0");
						oDS_PS_QM030H.SetValue("U_ChemC_Cu", 0, "0");
						oMat.Clear();

						sQry = "select a.ItemCode, left(a.DistNumber,8) from [OBTN] a ";
						sQry += "where left(a.DistNumber,8) = '" + oForm.Items.Item("LotNo").Specific.Value.ToString().Trim() + "' ";
						sQry += "group by a.ItemCode,left(a.DistNumber,8)";
						oRecordSet.DoQuery(sQry);

						while (!oRecordSet.EoF)
						{
							oForm.Items.Item("ItemCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oRecordSet.MoveNext();
						}
						break;

					case "ItemCode":
						sQry = "select ItemName from [OITM] where ItemCode = '" + oDS_PS_QM030H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_QM030H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						PS_QM030_Search_Matrix_Data();
						break;

					case "ChemC_Fe":
					case "ChemC_P":
					case "ChemC_Cu":
						oMat.FlushToDataSource();
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oDS_PS_QM030L.SetValue("U_ChemC_Fe", i, oDS_PS_QM030H.GetValue("U_ChemC_Fe", 0).ToString().Trim());
							oDS_PS_QM030L.SetValue("U_ChemC_P", i, oDS_PS_QM030H.GetValue("U_ChemC_P", 0).ToString().Trim());
							oDS_PS_QM030L.SetValue("U_ChemC_Cu", i, oDS_PS_QM030H.GetValue("U_ChemC_Cu", 0).ToString().Trim());
						}
						oMat.LoadFromDataSource();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM030_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM030_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_QM030H.GetValue("U_LotNo", 0).ToString().Trim()))
				{
					errMessage = "원소재 Lot No는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM030H.GetValue("U_ItemCode", 0).ToString().Trim()))
				{
					errMessage = "원소재 코드는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_QM030H.GetValue("U_ItemName", 0).ToString().Trim()))
				{
					errMessage = "원소재 이름이 없습니다. 원소재 코드를 확인하여 주십시오.";
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
		/// PS_QM030_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM030_MatrixSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하여 주십시오.";
					throw new Exception();
				}
				oMat.LoadFromDataSource();
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
		/// PS_QM030_Search_Matrix_Data
		/// </summary>
		private void PS_QM030_Search_Matrix_Data()
		{
			int Cnt;
			int j;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				oDS_PS_QM030H.SetValue("U_ChemC_Fe", 0, "0");
				oDS_PS_QM030H.SetValue("U_ChemC_P", 0, "0");
				oDS_PS_QM030H.SetValue("U_ChemC_Cu", 0, "0");

				sQry = "EXEC PS_QM030_01 '" + oDS_PS_QM030H.GetValue("U_LotNo", 0).ToString().Trim() + "', '" + oDS_PS_QM030H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				Cnt = oDS_PS_QM030L.Size;

				if (Cnt > 0)
				{
					for (j = 0; j <= Cnt - 1; j++)
					{
						oDS_PS_QM030L.RemoveRecord(oDS_PS_QM030L.Size - 1);
					}
					if (Cnt == 1)
					{
						oDS_PS_QM030L.Clear();
					}
				}
				oMat.LoadFromDataSource();

				j = 1;
				while (!oRecordSet.EoF)
				{
					if (oDS_PS_QM030L.Size < j)
					{
						oDS_PS_QM030L.InsertRecord(j - 1);
					}
					oDS_PS_QM030L.SetValue("LineId", j - 1, Convert.ToString(j));
					oDS_PS_QM030L.SetValue("U_CLotNo", j - 1, oRecordSet.Fields.Item("DistNumber").Value.ToString().Trim());
					oDS_PS_QM030L.SetValue("U_GRDate", j, Convert.ToDateTime(oRecordSet.Fields.Item("InDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_QM030L.SetValue("U_ChemC_Fe", j - 1, oRecordSet.Fields.Item("U_ChemC_Fe").Value.ToString().Trim());
					oDS_PS_QM030L.SetValue("U_ChemC_P", j - 1, oRecordSet.Fields.Item("U_ChemC_P").Value.ToString().Trim());
					oDS_PS_QM030L.SetValue("U_ChemC_Cu", j - 1, oRecordSet.Fields.Item("U_ChemC_Cu").Value.ToString().Trim());
					oDS_PS_QM030L.SetValue("U_Weight", j - 1, oRecordSet.Fields.Item("Quantity").Value.ToString().Trim());
					j += 1;
					oRecordSet.MoveNext();
				}
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM030_Update_OBTN
		/// </summary>
		private void PS_QM030_Update_OBTN()
		{
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oDS_PS_QM030L.Offset = i;

					sQry = "update [OBTN] set U_ChemC_Fe = '" + oDS_PS_QM030L.GetValue("U_ChemC_Fe", i).ToString().Trim() + "', ";
					sQry += "U_ChemC_P = '" + oDS_PS_QM030L.GetValue("U_ChemC_P", i).ToString().Trim() + "', ";
					sQry += "U_ChemC_Cu = '" + oDS_PS_QM030L.GetValue("U_ChemC_Cu", i).ToString().Trim() + "' ";
					sQry += "Where ItemCode = '" + oDS_PS_QM030H.GetValue("U_ItemCode", 0).ToString().Trim() + "' ";
					sQry += "and DistNumber = '" + oDS_PS_QM030L.GetValue("U_CLotNo", i).ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("성분 입력작업이 완료되었습니다", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM030_FormItemEnabled
		/// </summary>
		private void PS_QM030_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = true;
					oForm.Items.Item("ItemCode").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = true;
					oForm.Items.Item("ItemCode").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("LotNo").Enabled = false;
					oForm.Items.Item("ItemCode").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
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
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Register")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM030_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_QM030_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (Convert.ToDouble(oForm.Items.Item("ChemC_Fe").Specific.Value.ToString().Trim()) >= 0.05 && Convert.ToDouble(oForm.Items.Item("ChemC_Fe").Specific.Value.ToString().Trim()) <= 0.15)
							{
								if (Convert.ToDouble(oForm.Items.Item("ChemC_P").Specific.Value.ToString().Trim()) >= 0.025 && Convert.ToDouble(oForm.Items.Item("ChemC_P").Specific.Value.ToString().Trim()) <= 0.04)
								{
									PS_QM030_Update_OBTN();
									oForm.Items.Item("ChemC_Fe").Specific.Value = "0";
									oForm.Items.Item("ChemC_P").Specific.Value = "0";
									oForm.Items.Item("ChemC_Cu").Specific.Value = "0";

									oMat.Clear();
									oMat.FlushToDataSource();
									oMat.LoadFromDataSource();
									return;
								}
								errMessage = "P 의 값이 잘못 입력되었습니다. 확인해주세요.";
								throw new Exception();
							}
							errMessage = "Fe 의 값이 잘못 입력되었습니다. 확인해주세요.";
							throw new Exception();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_QM030_FormItemEnabled();
						}
					}
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
						if (pVal.ItemUID == "LotNo")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("LotNo").Specific.Value.ToString().Trim()))
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
						if (pVal.ItemUID == "ItemCode")
						{
							PS_QM030_FlushToItemValue(pVal.ItemUID);
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
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "LotNo" || pVal.ItemUID == "ChemC_Fe" || pVal.ItemUID == "ChemC_P" || pVal.ItemUID == "ChemC_Cu")
						{
							PS_QM030_FlushToItemValue(pVal.ItemUID);
						}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM030H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM030L);
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
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									oMat.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = i + 1;
								}

								oMat.FlushToDataSource();
								oDS_PS_QM030L.RemoveRecord(oDS_PS_QM030L.Size - 1);
								oMat.Clear();
								oMat.LoadFromDataSource();
							}
							break;
						case "1281": //찾기
							PS_QM030_FormItemEnabled();
							oForm.Items.Item("LotNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM030_FormItemEnabled();
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM030_FormItemEnabled();
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
