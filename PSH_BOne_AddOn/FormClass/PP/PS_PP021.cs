using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작업일보 작업자 분류등록
	/// </summary>
	internal class PS_PP021 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP021H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP021L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP021.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP021_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP021");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);

				PS_PP021_CreateItems();
				PS_PP021_SetComboBox();
				PS_PP021_SetDocument(oFormDocEntry);

				oForm.EnableMenu("1293", true); // 행삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1284", true); // 취소
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
			}
		}

		/// <summary>
		/// PS_PP021_CreateItems
		/// </summary>
		private void PS_PP021_CreateItems()
		{
			try
			{
				oDS_PS_PP021H = oForm.DataSources.DBDataSources.Item("@PS_PP021H");
				oDS_PS_PP021L = oForm.DataSources.DBDataSources.Item("@PS_PP021L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP021_SetComboBox
		/// </summary>
		private void PS_PP021_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE NOT IN('104','107') order by Code", "", false, false);

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
		/// PS_PP021_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP021_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP021_EnableFormItem();
					PS_PP021_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_PP021_EnableFormItem();
					oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP021_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP021_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				//행추가여부
				if (RowIserted == false)
				{
					oRow = oMat.RowCount;
					oDS_PS_PP021L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP021L.Offset = oRow;
				oDS_PS_PP021L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
				oDS_PS_PP021L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
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
		/// PS_PP021_EnableFormItem
		/// </summary>
		private void PS_PP021_EnableFormItem()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("OrdGbn").Enabled = true;
					oForm.Items.Item("Gubun").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("OrdGbn").Enabled = true;
					oForm.Items.Item("Gubun").Enabled = true;
					oForm.EnableMenu("1282", true); //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("OrdGbn").Enabled = false;
					oForm.Items.Item("Gubun").Enabled = false;
					oForm.EnableMenu("1282", true); //추가
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
		/// PS_PP021_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP021_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string OrdGbn;
			string Gubun;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				oMat.FlushToDataSource();

				switch (oUID)
				{
					case "Gubun":
						OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
						Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();

						sQry = "SELECT U_CdName From [@PS_SY001L] Where Code = 'P208' And U_RelCd = '" + OrdGbn + "' And U_Minor = '" + Gubun + "'";
						oRecordSet.DoQuery(sQry);
						break;

					case "Mat01":
						if (oCol == "CntcCode")
						{
							if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
							{
								oMat.FlushToDataSource();
								PS_PP021_AddMatrixRow(oMat.RowCount, false);
								oMat.Columns.Item("CntcCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}

							oMat.FlushToDataSource();

							sQry = "Select FULLNAME = t.U_FULLNAME, ";
							sQry += " TeamCode = t.U_TeamCode, ";
							sQry += " TeamNm = (Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code ";
							sQry += " Where a.Name = '부서' And b.U_Code = t.U_TeamCode ),  ";
							sQry += " RspCode = Isnull(t.U_RspCode,''), ";
							sQry += " RspNm = Isnull((Select U_CodeNm From [@PS_HR200H] a Inner Join [@PS_HR200L] b on a.Code = b.Code  ";
							sQry += " Where a.Name = '담당' And b.U_Code = t.U_RspCode ),'')  ";
							sQry += " From [@PH_PY001A] t Where U_status <> '5' And Code =  '" + oMat.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "' ";
							sQry += " And t.U_CLTCOD = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount > 0)
							{
								oDS_PS_PP021L.SetValue("U_FullName", oRow - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
								oDS_PS_PP021L.SetValue("U_TeamCode", oRow - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oDS_PS_PP021L.SetValue("U_TeamName", oRow - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_PP021L.SetValue("U_RspCode", oRow - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
								oDS_PS_PP021L.SetValue("U_RspName", oRow - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());

								oMat.Columns.Item("CntcCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oMat.LoadFromDataSource();
								oMat.FlushToDataSource();
								oMat.LoadFromDataSource();
							}
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Freeze(false);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP021_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP021_DelMatrixSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인데이타가 없습니다. 확인하세요.";
					throw new Exception();
				}
				else if (oMat.VisualRowCount == 1)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP021L.GetValue("U_CntcCode", 0).ToString().Trim()))
					{
						errMessage = "라인데이타가 없습니다. 확인하세요.";
						throw new Exception();
					}
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_PP021L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_PP021L.GetValue("U_CntcCode", i).ToString().Trim()))
						{
							errMessage = "사번은 필수입력사항입니다. 확인하세요.";
							throw new Exception();
						}
					}

					if (string.IsNullOrEmpty(oDS_PS_PP021L.GetValue("U_CntcCode", oMat.VisualRowCount - 1)))
					{
						oDS_PS_PP021L.RemoveRecord(oMat.VisualRowCount - 1);
					}
				}

				oMat.LoadFromDataSource();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP021_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP021_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_PP021H.GetValue("U_OrdGbn", 0).ToString().Trim()))
                {
					errMessage = "품목분류는 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP021H.GetValue("U_Gubun", 0).ToString().Trim()))
                {
					errMessage = "구분은 필수입력 사항입니다.";
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP021_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_PP021_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm.Items.Item("Code").Specific.Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() + oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP021_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_PP021_DelMatrixSpaceLine() == false)
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
						PS_PP021_EnableFormItem();
						PS_PP021_AddMatrixRow(0, true);
					}
					if (pVal.BeforeAction == false && pVal.ItemChanged == true)
					{
						if (pVal.ColUID == "CntcCode")
						{
							PS_PP021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Gubun")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("Gubun").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (string.IsNullOrEmpty(oMat.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ColUID == "CntcCode")
						{
							PS_PP021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							PS_PP021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "CntcCode")
							{
								PS_PP021_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					PS_PP021_AddMatrixRow(oMat.VisualRowCount, false);
					PS_PP021_EnableFormItem();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP021H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP021L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
							oForm.DataBrowser.BrowseBy = "Code";
							break;
						case "1282": //추가
							oForm.DataBrowser.BrowseBy = "Code";
							PS_PP021_AddMatrixRow(0, true);
							break;
						case "1285": //복원
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
						case "1281": //찾기
							PS_PP021_AddMatrixRow(0, true);
							PS_PP021_EnableFormItem();
							break;
						case "1282": //추가
							PS_PP021_EnableFormItem();
							PS_PP021_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							oDS_PS_PP021H.SetValue("Code", 0, "");

							for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								oMat.FlushToDataSource();
								oDS_PS_PP021L.SetValue("Code", i, "");
								oMat.LoadFromDataSource();
							}
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_PP021_EnableFormItem();
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 1; i <= oMat.VisualRowCount; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
								}
								oMat.FlushToDataSource(); // DBDataSource에 레코드가 한줄 더 생긴다.
								oDS_PS_PP021L.RemoveRecord(oDS_PS_PP021L.Size - 1); // 레코드 한 줄을 지운다.
								oMat.LoadFromDataSource(); // DBDataSource를 매트릭스에 올리고
								if (oMat.RowCount == 0)
								{
									PS_PP021_AddMatrixRow(1, false);
								}
								else
								{
									if (!string.IsNullOrEmpty(oDS_PS_PP021L.GetValue("U_CntcCode", oMat.RowCount - 1).ToString().Trim()))
									{
										PS_PP021_AddMatrixRow(oMat.RowCount, false);
									}
								}
							}
							break;
						case "7169": //엑셀 내보내기
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
				oForm.Freeze(false);
			}
		}
	}
}
