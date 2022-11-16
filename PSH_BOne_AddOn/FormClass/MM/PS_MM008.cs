using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 제품공정 단중 등록
	/// </summary>
	internal class PS_MM008 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_MM008H ;//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM008L; //등록라인

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM008.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM008_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM008");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				PS_MM008_CreateItems();
				PS_MM008_ComboBox_Setting();
				PS_MM008_Initialization();

				oForm.EnableMenu("1283", true);  // 삭제
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
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
		/// PS_MM008_CreateItems
		/// </summary>
		private void PS_MM008_CreateItems()
		{
			try
			{
				oDS_PS_MM008H = oForm.DataSources.DBDataSources.Item("@PS_MM008H");
				oDS_PS_MM008L = oForm.DataSources.DBDataSources.Item("@PS_MM008L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM008_ComboBox_Setting
		/// </summary>
		private void PS_MM008_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM008_Initialization
		/// </summary>
		private void PS_MM008_Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");
				PS_MM008_Add_MatrixRow(0, true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM008_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM008_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_MM008L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM008L.Offset = oRow;
				oDS_PS_MM008L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM008_Copy_MatrixRow
		/// </summary>
		private void PS_MM008_Copy_MatrixRow()
		{
			int i;

			try
			{
				oDS_PS_MM008H.SetValue("Code", 0, "");

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_MM008L.SetValue("Code", i, "");
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM008_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM008_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM008H.GetValue("U_ItmBsort", 0)))
				{
					errMessage = "대분류는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM008H.GetValue("U_BPLId", 0)))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM008H.GetValue("U_YM", 0)))
				{
					errMessage = "년월은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("YM").Specific.Value.ToString().Trim().Length != 6)
				{
					errMessage = "년월은 6자리 YYYYMM 형식으로 입력해야합니다. 확인하세요.";
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
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
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
<<<<<<< HEAD
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
=======
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
>>>>>>> C#Migration_MM
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
				//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
			string YM;
			string Code;
			string BPLID;
			string ItmBsort;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM008_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
							YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim().Substring(2, 4);
							ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
							Code = YM + BPLID + ItmBsort;
							oDS_PS_MM008H.SetValue("Code", 0, Code);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
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
				if (pVal.Before_Action == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "ItmBsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ItemCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							if (pVal.ColUID == "CpCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
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
			string ItemCode;
			string CpCode;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
						if (pVal.ItemUID == "ItmBsort")
						{
							sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItmBname").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "Mat01")
						{
							oMat.FlushToDataSource();
							if (pVal.ColUID == "BfWgt" || pVal.ColUID == "AfWgt")
							{
								if (Convert.ToDouble(oMat.Columns.Item("AfWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) != 0 || Convert.ToDouble(oMat.Columns.Item("BfWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) != 0)
								{
									oDS_PS_MM008L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim())));
									oDS_PS_MM008L.SetValue("U_Scrap", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat.Columns.Item("BfWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) - Convert.ToDouble(oMat.Columns.Item("AfWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))); //스크랩량
								}
								else
								{
									oDS_PS_MM008L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
								}

								oMat.LoadFromDataSource();
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
							}
							else if (pVal.ColUID == "ItemCode")
							{
								if ((pVal.Row == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									PS_MM008_Add_MatrixRow(oMat.RowCount, false);
									oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}

								ItemCode = oDS_PS_MM008L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
								sQry = "Select a.ItemCode, a.ItemName ";
								sQry += " From OITM a ";
								sQry += " Where a.ItemCode = '" + ItemCode + "'";
								oRecordSet.DoQuery(sQry);

								oMat.FlushToDataSource();

								if (oRecordSet.RecordCount == 0)
								{
									oDS_PS_MM008L.SetValue("U_ItemCode", pVal.Row - 1, "");
									oDS_PS_MM008L.SetValue("U_ItemName", pVal.Row - 1, "");
									oDS_PS_MM008L.SetValue("U_Cpcode", pVal.Row - 1, "");
									oDS_PS_MM008L.SetValue("U_CpName", pVal.Row - 1, "");
									oDS_PS_MM008L.SetValue("U_BfWgt", pVal.Row - 1, "0");
									oDS_PS_MM008L.SetValue("U_AfWgt", pVal.Row - 1, "0");
									oDS_PS_MM008L.SetValue("U_Scrap", pVal.Row - 1, "0");
									oDS_PS_MM008L.SetValue("U_Rate", pVal.Row - 1, "0");
									oMat.LoadFromDataSource();

									errMessage = "조회 결과가 없습니다. 확인하세요.";
									throw new Exception();
								}

								oDS_PS_MM008L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //품목코드
								oDS_PS_MM008L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품목이름
								oMat.LoadFromDataSource();
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
							}
							else if (pVal.ColUID == "CpCode")
							{
								CpCode = oDS_PS_MM008L.GetValue("U_CpCode", pVal.Row - 1).ToString().Trim();
								sQry = "Select U_CpName";
								sQry += " From [@PS_PP001L] ";
								sQry += " Where U_CpCode = '" + CpCode + "'";
								oRecordSet.DoQuery(sQry);
								oMat.FlushToDataSource();

								if (oRecordSet.RecordCount == 0)
								{
									oDS_PS_MM008L.SetValue("U_CpName", pVal.Row - 1, ""); //원재료명
								}
								else
								{
									oDS_PS_MM008L.SetValue("U_CpName", pVal.Row - 1, oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim()); //공정명
								}
								oMat.LoadFromDataSource();
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
							}
<<<<<<< HEAD
							else if (pVal.ColUID == "EdCpCode")
							{
								CpCode = oDS_PS_MM008L.GetValue("U_EdCpCode", pVal.Row - 1).ToString().Trim();
								sQry = "Select U_CpName";
								sQry += " From [@PS_PP001L] ";
								sQry += " Where U_CpCode = '" + CpCode + "'";
								oRecordSet.DoQuery(sQry);
								oMat.FlushToDataSource();

								if (oRecordSet.RecordCount == 0)
								{
									oDS_PS_MM008L.SetValue("U_EdCpName", pVal.Row - 1, ""); //원재료명
								}
								else
								{
									oDS_PS_MM008L.SetValue("U_EdCpName", pVal.Row - 1, oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim()); //공정명
								}
								oMat.LoadFromDataSource();
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
							}
							oMat.AutoResizeColumns();
=======
>>>>>>> C#Migration_MM
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
<<<<<<< HEAD
		/// Raise_EVENT_CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
=======
>>>>>>> C#Migration_MM
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
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				int i = 0;

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (oMat.RowCount != oMat.VisualRowCount)
					{
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
						}

						oMat.FlushToDataSource();
						oDS_PS_MM008L.RemoveRecord(oDS_PS_MM008L.Size - 1);
						oMat.Clear();
						oMat.LoadFromDataSource();
					}
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
							break;
						case "1281": //찾기
							PS_MM008_Initialization();
							break;
						case "1282": //추가
							PS_MM008_Initialization();
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1287": // 복제
							PS_MM008_Copy_MatrixRow();
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
