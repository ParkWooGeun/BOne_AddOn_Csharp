using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비가동시간등록(신)
	/// </summary>
	internal class PS_PP160 : PSH_BaseClass
	{
		private string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP160H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP160L; //등록라인
		private SAPbouiCOM.BoFormMode oLast_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP160.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP160_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP160");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();
				Initialization();
				FormClear();
				Add_MatrixRow(0, true);
				FormItemEnabled();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oDS_PS_PP160H = oForm.DataSources.DBDataSources.Item("@PS_PP160H");
				oDS_PS_PP160L = oForm.DataSources.DBDataSources.Item("@PS_PP160L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oDS_PS_PP160H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
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
		/// Initialization
		/// </summary>
		private void Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// FormClear
		/// </summary>
		private void FormClear()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP160'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oForm.Items.Item("DocNum").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocNum").Specific.Value = DocNum;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP160L.InsertRecord((oRow));
				}
				oMat.AddRow();
				oDS_PS_PP160L.Offset = oRow;
				oDS_PS_PP160L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// FormItemEnabled
		/// </summary>
		private void FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Part").Enabled = true;
					oForm.Items.Item("ItmBsort").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Part").Enabled = true;
					oForm.Items.Item("ItmBsort").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Part").Enabled = true;
					oForm.Items.Item("ItmBsort").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
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
		private void FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				switch (oUID)
				{
					case "ItmBsort":
						sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oDS_PS_PP160H.GetValue("U_ItmBsort", 0).ToString().Trim() +"'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_PP160H.SetValue("U_ItmBname", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						break;
					case "Mat01":
						switch (oCol)
						{
							case "MachCode":
								if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("MachCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									Add_MatrixRow(oMat.RowCount, false);
									oMat.Columns.Item("MachCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}

								sQry = "Select a.U_MachCode, a.U_MachName ";
								sQry += "From [@PS_PP130H] a ";
								sQry += "Where a.U_MachCode = '" + oMat.Columns.Item("MachCode").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
								oRecordSet.DoQuery(sQry);
								if (oRecordSet.RecordCount == 0)
								{
									errMessage = "조회 결과가 없습니다. 확인하세요.";
									throw new Exception();
								}
								oMat.Columns.Item("MachName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_MachName").Value.ToString().Trim();
								oMat.Columns.Item("MachCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								break;
							case "NCode01":
								sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" + oMat.Columns.Item("NCode01").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
								oRecordSet.DoQuery(sQry);
								oMat.Columns.Item("NCdName1").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
								break;
							case "NCode02":
								sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" + oMat.Columns.Item("NCode02").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
								oRecordSet.DoQuery(sQry);
								oMat.Columns.Item("NCdName2").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
								break;
							case "NCode03":
								sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" + oMat.Columns.Item("NCode03").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
								oRecordSet.DoQuery(sQry);
								oMat.Columns.Item("NCdName3").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
								break;
							case "NCode04":
								sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" + oMat.Columns.Item("NCode04").Cells.Item(oRow).Specific.Value.ToString().Trim() +"'";
								oRecordSet.DoQuery(sQry);
								oMat.Columns.Item("NCdName4").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
								break;
						}
						break;
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{ 
				if (string.IsNullOrEmpty(oDS_PS_PP160H.GetValue("U_BPLId", 0).ToString().Trim()))
                {
					errMessage = "사업장은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP160H.GetValue("U_Part", 0).ToString().Trim()))
                {
					errMessage = "담당은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP160H.GetValue("U_ItmBsort", 0).ToString().Trim()))
                {
					errMessage = "품목대분류는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP160H.GetValue("U_DocDate", 0).ToString().Trim()))
                {
					errMessage = "일자는 필수사항입니다. 확인하세요.";
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
		/// MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}
				else if (oMat.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_PP160L.GetValue("U_MachCode", 0).ToString().Trim()))
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}


				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP160L.GetValue("U_MachCode", i).ToString().Trim()))
					{
						errMessage = "설비코드는 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (Convert.ToDouble(oDS_PS_PP160L.GetValue("U_WorkTime", i).ToString().Trim()) == 0)
					{
						errMessage = "근무시간은 필수사항입니다. 확인하세요.";
						throw new Exception();
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
		/// Delete_EmptyRow
		/// </summary>
		private void Delete_EmptyRow()
		{
			int i;

			try
			{
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP160L.GetValue("U_MachCode", i).ToString().Trim()))
					{
						oDS_PS_PP160L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
					}
				}
				oMat.LoadFromDataSource();
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								oMat.FlushToDataSource();
								oMat.LoadFromDataSource();
							}
							Delete_EmptyRow();
							oLast_Mode = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
						{
							oLast_Mode = oForm.Mode;
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
							{
								Add_MatrixRow(oMat.RowCount, false);
								oLast_Mode = (BoFormMode)100;
							}
							else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
							{
								Add_MatrixRow(oMat.RowCount, false);
								FormItemEnabled();
								oLast_Mode = (BoFormMode)100;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
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
						if (pVal.ItemUID == "Part")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("Part").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ItmBsort")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "Mat01")
						{
							switch (pVal.ColUID)
							{
								case "MachCode":
									if (string.IsNullOrEmpty(oMat.Columns.Item("MachCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
									break;
								case "NCode01":
									if (string.IsNullOrEmpty(oMat.Columns.Item("NCode01").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
									break;
								case "NCode02":
									if (string.IsNullOrEmpty(oMat.Columns.Item("NCode02").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
									break;
								case "NCode03":
									if (string.IsNullOrEmpty(oMat.Columns.Item("NCode03").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
									break;
								case "NCode04":
									if (string.IsNullOrEmpty(oMat.Columns.Item("NCode04").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
									break;
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
					if (pVal.ItemUID == "BPLId")
					{
						oMat.Clear();
						oDS_PS_PP160L.Clear();
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							Add_MatrixRow(0, false);
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "ItmBsort")
						{
							FlushToItemValue(pVal.ItemUID, 0 ,"");
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "MachCode")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "NCode01")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "NCode02")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "NCode03")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "NCode04")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP160H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP160L);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
							FormItemEnabled();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1286": //닫기
							break;
						case "1281": //찾기
							FormItemEnabled();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							break;
						case "1282": //추가
							Initialization();
							FormItemEnabled();
							FormClear();
							oDS_PS_PP160H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							Add_MatrixRow(0, true);
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							FormItemEnabled();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("MachCode").Cells.Item(oMat.VisualRowCount).Specific.Value.ToString().Trim()))
								{
									if (oDS_PS_PP160H.GetValue("Status", 0) == "O")
									{
										Add_MatrixRow(oMat.RowCount, false);
									}
								}
							}
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
								}
								oMat.FlushToDataSource();
								oDS_PS_PP160L.RemoveRecord(oDS_PS_PP160L.Size - 1); // Mat01에 마지막라인(빈라인) 삭제
								oMat.Clear();
								oMat.LoadFromDataSource();
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
		}
	}
}
