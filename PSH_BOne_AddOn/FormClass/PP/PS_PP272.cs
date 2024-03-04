using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 계량기별전력사용량등록
	/// </summary>
	internal class PS_PP272 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP272H; //등록헤더 
		private SAPbouiCOM.DBDataSource oDS_PS_PP272L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP272.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP272_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP272");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_PP272_CreateItems();
				PS_PP272_SetComboBox();
				PS_PP272_EnableFormItem();
				PS_PP272_ClearForm();
				PS_PP272_AddMatrixRow(0, oMat.RowCount, true);

				oForm.EnableMenu("1283", true);  // 제거
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1284", false); // 취소
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
		/// PS_PP272_CreateItems
		/// </summary>
		private void PS_PP272_CreateItems()
		{
			try
			{
				oDS_PS_PP272H = oForm.DataSources.DBDataSources.Item("@PS_PP272H");
				oDS_PS_PP272L = oForm.DataSources.DBDataSources.Item("@PS_PP272L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP272_SetComboBox
		/// </summary>
		private void PS_PP272_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_PP272_EnableFormItem
		/// </summary>
		private void PS_PP272_EnableFormItem()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("YM").Enabled = true;
					oForm.Items.Item("TeamCode").Enabled = true;
					oForm.Items.Item("RspCode").Enabled = true;
					oForm.Items.Item("Btn_Set").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("YM").Enabled = true;
					oForm.Items.Item("TeamCode").Enabled = true;
					oForm.Items.Item("RspCode").Enabled = true;
					oForm.Items.Item("Btn_Set").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("YM").Enabled = false;
					oForm.Items.Item("TeamCode").Enabled = false;
					oForm.Items.Item("RspCode").Enabled = false;
					oForm.Items.Item("Btn_Set").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP272_ClearForm
		/// </summary>
		private void PS_PP272_ClearForm()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP272'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oDS_PS_PP272H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_PP272H.SetValue("DocEntry", 0, DocNum); // 화면에 적용이 안되기 때문
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP272_AddMatrixRow
		/// </summary>
		/// <param name="oSeq"></param>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP272_AddMatrixRow(int oSeq, int oRow, bool RowIserted)
		{
			try
			{
				switch (oSeq)
				{
					case 0:
						oDS_PS_PP272L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
					case 1:
						oDS_PS_PP272L.InsertRecord(oRow);
						oDS_PS_PP272L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP272_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP272_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
				{
					errMessage = "년월은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "부서코드는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_PP272_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP272_DelMatrixSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			
			try
			{
				oMat.FlushToDataSource();
				
				if (oMat.VisualRowCount <= 1)  // 라인
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}

				if (oMat.VisualRowCount > 0)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP272L.GetValue("U_Gauge", oMat.VisualRowCount - 1)))
					{
						oDS_PS_PP272L.RemoveRecord(oMat.VisualRowCount - 1);
					}
				}
				oMat.LoadFromDataSource();
				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_PP272_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oCID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP272_FlushToItemValue(string oUID, string oCID, int oRow, string oCol)
		{
			try
			{
				switch (oUID)
				{
					case "Mat01":
						switch (oCID)
						{
							case "CUsed":
								if ((oRow == oMat.RowCount || oMat.VisualRowCount == 2) && !string.IsNullOrEmpty(oMat.Columns.Item("CUsed").Cells.Item(oRow).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									PS_PP272_AddMatrixRow(1, oMat.RowCount, true);
									oMat.Columns.Item("CUsed").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
								break;

							case "AUsed":
								if ((oRow == oMat.RowCount || oMat.VisualRowCount == 2) && !string.IsNullOrEmpty(oMat.Columns.Item("AUsed").Cells.Item(oRow).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									PS_PP272_AddMatrixRow(1, oMat.RowCount, true);
									oMat.Columns.Item("AUsed").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
								break;
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP272_DataSet
		/// 기본자료 Load
		/// </summary>
		private void PS_PP272_DataSet()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string BPLId;
			string YM;
			string LYM;
			string TeamCode; 
			string RspCode;
			DateTime LDate;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();

				LDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("YM").Specific.Value.ToString().Trim() + "01", "-"));
				LYM = Convert.ToString(LDate.AddMonths(-1).ToString("yyyyMMdd")).Substring(0,6);  //전월

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "SELECT b.U_Gauge, ";
				sQry += "      b.U_GName, ";
				sQry += "	   b.U_RValue, ";
				sQry += "	   LUsed = Isnull((SELECT isnull(d.U_CUsed, 0) ";
				sQry += "						 FROM [@PS_PP272H] c INNER JOIN[@PS_PP272L] d ON c.DocEntry = d.DocEntry AND c.Canceled = 'N' ";
				sQry += "						  WHERE c.U_BPLId = '" + BPLId + "'";
				sQry += "						  AND c.U_YM = '" + LYM + "'";
				sQry += "						  AND c.U_TeamCode = '" + TeamCode + "'";
				sQry += "						  AND isnull(c.U_RspCode,'') = '" + RspCode + "'";
				sQry += "						  AND d.U_Gauge = b.U_Gauge),0) ";
				sQry += "  FROM[@PS_PP270H] a INNER JOIN[@PS_PP270L] b ON a.DocEntry = b.DocEntry AND a.Canceled = 'N' ";
				sQry += " WHERE a.U_BPLId = '" + BPLId + "'";
				sQry += "   AND a.U_TeamCode = '" + TeamCode + "'";
				sQry += "   AND isnull(a.U_RspCode,'') = '" + RspCode + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP272L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "계량기코드자료가 없습니다. 확인하세요.";
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP272_AddMatrixRow(0, oMat.RowCount, true);
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP272L.Size)
					{
						oDS_PS_PP272L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP272L.Offset = i;

					oDS_PS_PP272L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP272L.SetValue("U_Gauge", i,  oRecordSet.Fields.Item("U_Gauge").Value.ToString().Trim());  //계량기코드
					oDS_PS_PP272L.SetValue("U_GName", i,  oRecordSet.Fields.Item("U_GName").Value.ToString().Trim());  //계량기명
					oDS_PS_PP272L.SetValue("U_RValue", i, oRecordSet.Fields.Item("U_RValue").Value.ToString().Trim()); //보정값
					oDS_PS_PP272L.SetValue("U_LUsed", i,  oRecordSet.Fields.Item("LUsed").Value.ToString().Trim());    //전월지침량

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
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
			string BPLId;
			string YM;
			string TeamCode;
			string RspCode;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
	
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP272_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_PP272_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
					else if (pVal.ItemUID == "Btn_Set")
					{
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                            YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
							TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
							RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
							sQry = "SELECT count(*) From [@PS_PP272H] WHERE U_BPLId = '" + BPLId + "' AND U_YM = '" + YM + "' AND U_TeamCode = '" + TeamCode + "' AND U_RspCode = '" + RspCode + "' ";
							oRecordSet.DoQuery(sQry);

                            if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
                            {
                                PSH_Globals.SBO_Application.MessageBox("이미 등록된 자료 입니다 조회후 추가 하십시요..");
                                BubbleEvent = false;
                                return;
                            }
                        }

						PS_PP272_DataSet();
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
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == false)
						{
							PS_PP272_EnableFormItem();
							PS_PP272_AddMatrixRow(1, oMat.RowCount, true);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					if (string.IsNullOrEmpty(oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim()))
					{
						if (pVal.ItemUID == "TeamCode" && pVal.CharPressed == 9)  //부서코드  (담당은 없을수도있음)
						{
							oForm.Items.Item("TeamCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
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
			Double LUsed;
			Double CUsed;
			Double AUsed;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						PS_PP272_FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row, "");

						if (pVal.ItemUID == "TeamCode")
						{
							sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Code = '" + oForm.Items.Item("TeamCode").Specific.Value.Trim() + "' AND U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value.Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "RspCode")
						{
							sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Code = '" + oForm.Items.Item("RspCode").Specific.Value.Trim() + "' AND U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value.Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "CUsed")
							{
								oMat.FlushToDataSource();

								//사용량 = 당월지침 - 전월지침 * 보정값
								//LUsed = Convert.ToDouble(oDS_PS_PP272L.GetValue("U_LUsed", oMat.RowCount - 1).ToString().Trim());
								LUsed = Convert.ToDouble(oMat.Columns.Item("LUsed").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								CUsed = Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								AUsed = System.Math.Round((CUsed - LUsed) * Convert.ToDouble(oMat.Columns.Item("RValue").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()), 2);
								oDS_PS_PP272L.SetValue("U_AUsed", pVal.Row - 1, Convert.ToString(AUsed));

								oMat.LoadFromDataSource();
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
					}
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
					PS_PP272_AddMatrixRow(1, oMat.VisualRowCount, true);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP272H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP272L);
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
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
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
							PS_PP272_EnableFormItem();
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
							break;
						case "1282": //추가
							PS_PP272_EnableFormItem();
							PS_PP272_ClearForm();
							PS_PP272_AddMatrixRow(0, oMat.RowCount, true);
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
							oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							PS_PP272_EnableFormItem();
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 1; i <= oMat.VisualRowCount; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
								}
								oMat.FlushToDataSource(); // DBDataSource에 레코드가 한줄 더 생긴다.
								oDS_PS_PP272L.RemoveRecord(oDS_PS_PP272L.Size - 1);	// 레코드 한 줄을 지운다.
								oMat.LoadFromDataSource(); // DBDataSource를 매트릭스에 올리고
								if (oMat.RowCount == 0)
								{
									PS_PP272_AddMatrixRow(1, 0, true);
								}
								else
								{
									if (!string.IsNullOrEmpty(oDS_PS_PP272L.GetValue("U_Gauge", oMat.RowCount - 1).ToString().Trim()))
									{
										PS_PP272_AddMatrixRow(1, oMat.RowCount, true);
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
