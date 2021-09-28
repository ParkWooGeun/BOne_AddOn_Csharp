using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월생산계획대비실적조회
	/// </summary>
	internal class PS_PP860 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP860H;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP860.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_PP860_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP860");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

				oForm.Freeze(true);

				PS_PP860_CreateItems();
				PS_PP860_ComboBox_Setting();
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
		/// PS_PP860_CreateItems
		/// </summary>
		private void PS_PP860_CreateItems()
		{
			try
			{
				oDS_PS_PP860H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
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
		/// PS_PP860_ComboBox_Setting
		/// </summary>
		private void PS_PP860_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("YYYYMM").Specific.Value = DateTime.Now.ToString("yyyyMM");

				oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);

				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("YYYYMM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP860_MatrixColumnSetting
		/// </summary>
		/// <param name="pYM"></param>
		private void PS_PP860_MatrixColumnSetting(string pYM)
		{
			int loopCount;
			DateTime Ymd;
			string Dt;
			int LastDay;
			int DisableColumn;
			string DisableColumnString;

			string DayName = string.Empty;
			string temp;

			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				Ymd = Convert.ToDateTime(dataHelpClass.ConvertDateType(pYM + "01", "-"));
				Dt = Convert.ToString(Ymd.AddMonths(1).AddDays(-1).ToString("yyyyMMdd")); //해당 월의 마지막 날
				LastDay = Convert.ToInt32(codeHelpClass.Right(Dt, 2));

				//LineNum는 제외하기 위해서 1부터 시작
				for (loopCount = 1; loopCount <= oMat.Columns.Count - 1; loopCount++)
				{
					oMat.Columns.Item(loopCount).Editable = true;
				}

				for (loopCount = 1; loopCount <= LastDay; loopCount++)
				{
					if (loopCount < 10)
					{
						temp = "0";
					}
					else
					{
						temp = "";
					}

					switch (Convert.ToDateTime(dataHelpClass.ConvertDateType(codeHelpClass.Left(pYM, 6) + temp + loopCount.ToString(), "-")).DayOfWeek)
					//switch (dd)
					{
						case DayOfWeek.Sunday:
							DayName = "일";
							break;
						case DayOfWeek.Monday:
							DayName = "월";
							break;
						case DayOfWeek.Tuesday:
							DayName = "화";
							break;
						case DayOfWeek.Wednesday:
							DayName = "수";
							break;
						case DayOfWeek.Thursday:
							DayName = "목";
							break;
						case DayOfWeek.Friday:
							DayName = "금";
							break;
						case DayOfWeek.Saturday:
							DayName = "토";
							break;
					}

					DisableColumnString = "D" + temp + Convert.ToString(loopCount);
					oMat.Columns.Item(DisableColumnString).TitleObject.Caption = Convert.ToString(loopCount) + "일(" + DayName + ")";

					if (DayName == "일")
					{
						oMat.Columns.Item(DisableColumnString).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 0, 0)); //빨간색
					}
					else if (DayName == "토")
					{
						oMat.Columns.Item(DisableColumnString).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 128, 255)); //하늘색
					}
					else
					{
						oMat.Columns.Item(DisableColumnString).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));//흰색
					}
				}

				if (LastDay != 31)
				{
					DisableColumn = 31 - LastDay;

					for (loopCount = 0; loopCount <= DisableColumn - 1; loopCount++)
					{
						DisableColumnString = "D" + Convert.ToString(31 - loopCount);
						oMat.Columns.Item(DisableColumnString).Editable = false; //해당월의 말일이 존재하지 않으면 막음
					}
				}
				oMat.AutoResizeColumns();
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
		/// PS_PP860_LoadData 조회데이타 가져오기
		/// </summary>
		private void PS_PP860_LoadData()
		{
			short i;
			string sQry;

			string YYYYMM;
			string BPLId;
			string Part;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
				Part = oForm.Items.Item("Part").Specific.Value.ToString().Trim();
				
				if (codeHelpClass.Left(Part, 3) == "124")  //부품
				{
					sQry = "EXEC [PS_PP860_02] '" + BPLId + "','" + YYYYMM + "','" + Part + "'";
				}
				else
				{
					sQry = "EXEC [PS_PP860_01] '" + BPLId + "','" + YYYYMM + "','" + Part + "'";
				}

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP860H.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
					if (i + 1 > oDS_PS_PP860H.Size)
					{
						oDS_PS_PP860H.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP860H.Offset = i;

					oDS_PS_PP860H.SetValue("U_ColRgl01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColRgl02", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColRgl03", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColRgl04", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(10).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg08", i, oRecordSet.Fields.Item(11).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg09", i, oRecordSet.Fields.Item(12).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg10", i, oRecordSet.Fields.Item(13).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg11", i, oRecordSet.Fields.Item(14).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg12", i, oRecordSet.Fields.Item(15).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg13", i, oRecordSet.Fields.Item(16).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg14", i, oRecordSet.Fields.Item(17).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg15", i, oRecordSet.Fields.Item(18).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg16", i, oRecordSet.Fields.Item(19).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg17", i, oRecordSet.Fields.Item(20).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg18", i, oRecordSet.Fields.Item(21).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg19", i, oRecordSet.Fields.Item(22).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg20", i, oRecordSet.Fields.Item(23).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg21", i, oRecordSet.Fields.Item(24).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg22", i, oRecordSet.Fields.Item(25).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg23", i, oRecordSet.Fields.Item(26).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg24", i, oRecordSet.Fields.Item(27).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg25", i, oRecordSet.Fields.Item(28).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg26", i, oRecordSet.Fields.Item(29).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg27", i, oRecordSet.Fields.Item(30).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg28", i, oRecordSet.Fields.Item(31).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg29", i, oRecordSet.Fields.Item(32).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg30", i, oRecordSet.Fields.Item(33).Value.ToString().Trim());
					oDS_PS_PP860H.SetValue("U_ColReg31", i, oRecordSet.Fields.Item(34).Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
			    }

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
		    }
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP860_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP860_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP860H.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP860H.Offset = oRow;
				oDS_PS_PP860H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP860_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP860_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					errMessage = "사업장은 필수사항입니다. 확인하세요.:";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("YYYYMM").Specific.ToString().Trim()))
				{
					errMessage = "년월은 필수사항입니다. 확인하세요.:";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("Part").Specific.Value.ToString().Trim()))
				{
					errMessage = "담당은 필수사항입니다. 확인하세요.:";
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "Btn_ret")
					{
						if (PS_PP860_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						PS_PP860_MatrixColumnSetting(oForm.Items.Item("YYYYMM").Specific.Value + "01");
						PS_PP860_LoadData();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP860H);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
