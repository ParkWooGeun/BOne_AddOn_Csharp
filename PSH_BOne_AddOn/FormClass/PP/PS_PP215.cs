using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 원재료 월소요량 등록
	/// </summary>
	internal class PS_PP215 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP215H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP215L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP215.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP215_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP215");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();
				Initialization();

				oForm.EnableMenu(("1283"), true);  // 삭제
				oForm.EnableMenu(("1287"), true);  // 복제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1284"), false); // 취소
				oForm.EnableMenu(("1293"), true);  // 행삭제
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
				oDS_PS_PP215H = oForm.DataSources.DBDataSources.Item("@PS_PP215H");
				oDS_PS_PP215L = oForm.DataSources.DBDataSources.Item("@PS_PP215L");
				oMat = oForm.Items.Item("Mat01").Specific;

				oForm.DataSources.UserDataSources.Add("S_Weight1", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("S_Weight1").Specific.DataBind.SetBound(true, "", "S_Weight1");

				oForm.DataSources.UserDataSources.Add("S_Weight2", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("S_Weight2").Specific.DataBind.SetBound(true, "", "S_Weight2");

				oForm.DataSources.UserDataSources.Add("S_PWeight1", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("S_PWeight1").Specific.DataBind.SetBound(true, "", "S_PWeight1");

				oForm.DataSources.UserDataSources.Add("S_PWeight2", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("S_PWeight2").Specific.DataBind.SetBound(true, "", "S_PWeight2");
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
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				//매트릭스 거래처
				sQry = "SELECT T1.CardCode, T1.CardName ";
				sQry += " FROM OCRD T1";
				sQry += " WHERE T1.CardType = 'C' And T1.frozenFor <> 'Y'";
				sQry += " and exists ( Select * From [@PS_SY001L] a ";
				sQry += " Where a.Code= 'S005' and a.U_UseYN = 'Y'";
				sQry += " and a.U_Minor = T1.CardCode)";
				sQry += " Order by T1.CardFName";
				oRecordSet.DoQuery(sQry);

				while (!(oRecordSet.EoF))
				{
					oMat.Columns.Item("CardCode").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
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

				oForm.Items.Item("YM").Specific.Value = DateTime.Now.ToString("yyyyMM");
				Add_MatrixRow(0, true);
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
					oDS_PS_PP215L.InsertRecord((oRow));
				}
				oMat.AddRow();
				oDS_PS_PP215L.Offset = oRow;
				oDS_PS_PP215L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (string.IsNullOrEmpty(oDS_PS_PP215H.GetValue("U_ItmBsort", 0).ToString().Trim()))
				{
					errMessage = "대분류는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP215H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP215H.GetValue("U_YM", 0).ToString().Trim()))
				{
					errMessage = "년월은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("YM").Specific.Value.ToString().Trim().Length != 6)
				{
					errMessage = "년월은 6자리 YYYYMM 형식으로 입력해야합니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP215H.GetValue("U_Rate", 0).ToString().Trim()))
				{
					errMessage = "기준수율은 필수입력사항입니다. 확인하세요.";
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
			string YM;
			string Code;
			string BPLId;
			string ItmBsort;
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

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

							BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
							YM = codeHelpClass.Right(oForm.Items.Item("YM").Specific.Value.ToString().Trim(), 4);
							ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
							Code = YM + BPLId + ItmBsort;
							oDS_PS_PP215H.SetValue("Code", 0, Code);
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
							if (pVal.ColUID == "PItemCod")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("PItemCod").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
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
			string sQry;
			string ItemCode;
			string errMessage = string.Empty;

			double rate_Renamed;
			double StdWgt;
			double S_StdWgt = 0;
			double S_Weight1 = 0;
			double S_Weight2 = 0;
			double S_PWeight1 = 0;
			double S_PWeight2 = 0;
			double S_AWeight1 = 0;

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
							sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItmBname").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "Mat01")
						{
							oMat.FlushToDataSource();
							if (pVal.ColUID == "StdWgt")
							{
								rate_Renamed = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());

								if (rate_Renamed != 0)
								{
									//기준수량변경시 초기화
									oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oDS_PS_PP215L.SetValue("U_Weight1", pVal.Row - 1, oMat.Columns.Item("StdWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //창원생산량
									oDS_PS_PP215L.SetValue("U_Weight2", pVal.Row - 1, "0");	//부산생산량
									oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) / rate_Renamed, 0))); //창원소요량
									oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, "0"); //부산소요량
								}
								else
								{
									oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
									PSH_Globals.SBO_Application.MessageBox("기준수율을 입력해야합니다.");
								}
								oMat.LoadFromDataSource();

								for (i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									S_StdWgt += Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight1 += Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight2 += Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight1 += Convert.ToDouble(oMat.Columns.Item("PWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight2 += Convert.ToDouble(oMat.Columns.Item("PWeight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
								}

								oForm.Items.Item("S_StdWgt").Specific.Value = S_StdWgt;
								oForm.Items.Item("S_Weight1").Specific.Value = S_Weight1;
								oForm.Items.Item("S_Weight2").Specific.Value = S_Weight2;
								oForm.Items.Item("S_PWeight1").Specific.Value = S_PWeight1;
								oForm.Items.Item("S_PWeight2").Specific.Value = S_PWeight2;
							}
							else if (pVal.ColUID == "Weight1")
							{
								//창원생산중량 입력
								StdWgt = Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								rate_Renamed = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());

								if (rate_Renamed != 0)
								{
									if (StdWgt > 0)
									{
										//창원원재료 소요량
										oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
										oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) / rate_Renamed, 0))); //창원소요량

										//기계중량은 기준중량 - 창원생산량
										oDS_PS_PP215L.SetValue("U_Weight2", pVal.Row - 1, Convert.ToString(StdWgt - Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())));	//부산생산량
										oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(StdWgt - Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) / rate_Renamed, 0))); //부산소요량
									}
									else
									{
										oDS_PS_PP215L.SetValue("U_Weight2", pVal.Row - 1, "0");	 //부산생산량
										oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, "0"); //창원소요량
										oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, "0"); //부산소요량
									}
									oMat.LoadFromDataSource();
								}
								else
								{
									oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
									PSH_Globals.SBO_Application.MessageBox("기준수율을 입력해야합니다.");
								}

								for (i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									S_StdWgt += Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight1 += Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight2 += Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight1 += Convert.ToDouble(oMat.Columns.Item("PWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight2 += Convert.ToDouble(oMat.Columns.Item("PWeight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
								}
								oForm.Items.Item("S_StdWgt").Specific.Value = S_StdWgt;
								oForm.Items.Item("S_Weight1").Specific.Value = S_Weight1;
								oForm.Items.Item("S_Weight2").Specific.Value = S_Weight2;
								oForm.Items.Item("S_PWeight1").Specific.Value = S_PWeight1;
								oForm.Items.Item("S_PWeight2").Specific.Value = S_PWeight2;

								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
							}
							else if (pVal.ColUID == "AWeight1")
							{
								rate_Renamed = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());

								if (rate_Renamed != 0)
								{
									//기준수량변경시 초기화
									oDS_PS_PP215L.SetValue("U_StdWgt", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) + Convert.ToDouble(oMat.Columns.Item("AWeight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())));
                                    oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, Convert.ToString(System.Math.Round((Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) + Convert.ToDouble(oMat.Columns.Item("AWeight1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) / rate_Renamed, 0))); //창원소요량
                                }
                                else
								{
									oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
									PSH_Globals.SBO_Application.MessageBox("기준수율을 입력해야합니다.");
								}
								oMat.LoadFromDataSource();

								for (i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									S_StdWgt += Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(i + 1).Specific.Value.ToString().Trim()) + Convert.ToDouble(oMat.Columns.Item("AWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight1 += Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight1 += Convert.ToDouble(oMat.Columns.Item("PWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_AWeight1 += Convert.ToDouble(oMat.Columns.Item("AWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
								}
								oForm.Items.Item("S_StdWgt").Specific.Value = S_StdWgt;
								oForm.Items.Item("S_Weight1").Specific.Value = S_Weight1 + S_AWeight1;
								oForm.Items.Item("S_PWeight1").Specific.Value = S_PWeight1;
							}
							else if (pVal.ColUID == "Weight2")
							{
								//부산생산중량 입력
								StdWgt = Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								rate_Renamed = Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim());

								if (rate_Renamed != 0)
								{
									if (StdWgt > 0)
									{
										oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) / rate_Renamed, 0))); //부산소요량

										//창원중량은 기준중량 - 부산생산량
										oDS_PS_PP215L.SetValue("U_Weight1", pVal.Row - 1, Convert.ToString(StdWgt - Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())));
										oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, Convert.ToString(System.Math.Round((StdWgt - Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())) / rate_Renamed, 0)));
									}
									else
									{
										oDS_PS_PP215L.SetValue("U_Weight1", pVal.Row - 1, "0");	 //창원생산량
										oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, "0"); //창원소요량
										oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, "0"); //부산소요량
									}
									oMat.LoadFromDataSource();
								}
								else
								{
									oDS_PS_PP215L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
									PSH_Globals.SBO_Application.MessageBox("기준수율을 입력해야합니다.");
								}
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

								for (i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									S_StdWgt += Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight1 += Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_Weight2 += Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight1 += Convert.ToDouble(oMat.Columns.Item("PWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
									S_PWeight2 += Convert.ToDouble(oMat.Columns.Item("PWeight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
								}
								oForm.Items.Item("S_StdWgt").Specific.Value = S_StdWgt;
								oForm.Items.Item("S_Weight1").Specific.Value = S_Weight1;
								oForm.Items.Item("S_Weight2").Specific.Value = S_Weight2;
								oForm.Items.Item("S_PWeight1").Specific.Value = S_PWeight1;
								oForm.Items.Item("S_PWeight2").Specific.Value = S_PWeight2;
							}
							else if (pVal.ColUID == "ItemCode")
							{
								if ((pVal.Row == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									Add_MatrixRow(oMat.RowCount, false);
									oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}

								ItemCode = oDS_PS_PP215L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
								sQry = "Select a.U_ItemCod1, a.U_ItemNam1, b.U_OutSize, b.U_CallSize, a.U_ItemCod2, a.U_ItemNam2";
								sQry += " From [@PS_PP005H] a Inner Join OITM b On a.U_ItemCod1 = b.ItemCode ";
								sQry += " Where a.U_ItemCod1 = '" + ItemCode + "'";
								oRecordSet.DoQuery(sQry);

								oMat.FlushToDataSource();

								if (oRecordSet.RecordCount == 0)
                                {
									//매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
									oDS_PS_PP215L.SetValue("U_ItemCode", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_ItemName", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_OutSize", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_CallSize", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_PItemCod", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_PItemNam", pVal.Row - 1, "");
									oDS_PS_PP215L.SetValue("U_StdWgt", pVal.Row - 1, "0");
									oDS_PS_PP215L.SetValue("U_Weight1", pVal.Row - 1, "0");
									oDS_PS_PP215L.SetValue("U_Weight2", pVal.Row - 1, "0");
									oDS_PS_PP215L.SetValue("U_PWeight1", pVal.Row - 1, "0");
									oDS_PS_PP215L.SetValue("U_PWeight2", pVal.Row - 1, "0");

									oMat.LoadFromDataSource();

									errMessage = "조회 결과가 없습니다. 확인하세요.";
									throw new Exception();
								}
								//매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
								oDS_PS_PP215L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item("U_ItemCod1").Value.ToString().Trim()); //품목코드
								oDS_PS_PP215L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("U_ItemNam1").Value.ToString().Trim()); //품목이름
								oDS_PS_PP215L.SetValue("U_OutSize", pVal.Row - 1, oRecordSet.Fields.Item("U_OutSize").Value.ToString().Trim());	  //외부규격
								oDS_PS_PP215L.SetValue("U_CallSize", pVal.Row - 1, oRecordSet.Fields.Item("U_CallSize").Value.ToString().Trim()); //호칭규격
								oDS_PS_PP215L.SetValue("U_PItemCod", pVal.Row - 1, oRecordSet.Fields.Item("U_ItemCod2").Value.ToString().Trim()); //원재료코드
								oDS_PS_PP215L.SetValue("U_PItemNam", pVal.Row - 1, oRecordSet.Fields.Item("U_ItemNam2").Value.ToString().Trim()); //원재료명

								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "PItemCod")
							{
								ItemCode = oDS_PS_PP215L.GetValue("U_PItemCod", pVal.Row - 1).ToString().Trim();
								sQry = "Select ItemName";
								sQry += " From OITM ";
								sQry += " Where ItemCode = '" + ItemCode + "'";
								oRecordSet.DoQuery(sQry);

								oMat.FlushToDataSource();

								if (oRecordSet.RecordCount == 0)
								{
									oDS_PS_PP215L.SetValue("U_PItemNam", pVal.Row - 1, "");	//원재료명
								}
								else
								{
									oDS_PS_PP215L.SetValue("U_PItemNam", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());	//원재료명
								}

								oMat.LoadFromDataSource();
							}
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
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			double S_StdWgt = 0;
			double S_Weight1 = 0;
			double S_Weight2 = 0;
			double S_PWeight1 = 0;
			double S_PWeight2 = 0;
			double S_AWeight1 = 0;

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						S_StdWgt += Convert.ToDouble(oMat.Columns.Item("StdWgt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
						S_Weight1 += Convert.ToDouble(oMat.Columns.Item("Weight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
						S_Weight2 += Convert.ToDouble(oMat.Columns.Item("Weight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
						S_PWeight1 += Convert.ToDouble(oMat.Columns.Item("PWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
						S_PWeight2 += Convert.ToDouble(oMat.Columns.Item("PWeight2").Cells.Item(i + 1).Specific.Value.ToString().Trim());
						S_AWeight1 += Convert.ToDouble(oMat.Columns.Item("AWeight1").Cells.Item(i + 1).Specific.Value.ToString().Trim());
					}
					oForm.Items.Item("S_StdWgt").Specific.Value = S_StdWgt;
					oForm.Items.Item("S_Weight1").Specific.Value = S_Weight1 + S_AWeight1;
					oForm.Items.Item("S_Weight2").Specific.Value = S_Weight2;
					oForm.Items.Item("S_PWeight1").Specific.Value = S_PWeight1;
					oForm.Items.Item("S_PWeight2").Specific.Value = S_PWeight2;
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP215H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP215L);
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
							break;
						case "1282": //추가
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
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
							Initialization();
							break;
						case "1282": //추가
							Initialization();
							break;
						case "1287": //복제
							oDS_PS_PP215H.SetValue("Code", 0, "");

							for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								oMat.FlushToDataSource();
								oDS_PS_PP215L.SetValue("Code", i, "");
								oMat.LoadFromDataSource();
							}
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
								}

								oMat.FlushToDataSource();
								oDS_PS_PP215L.RemoveRecord(oDS_PS_PP215L.Size - 1); // Mat01에 마지막라인(빈라인) 삭제
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
