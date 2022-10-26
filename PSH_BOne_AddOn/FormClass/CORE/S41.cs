using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 배치번호추가
	/// </summary>
	internal class S41 : PSH_BaseClass
	{
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;

		private int oMatTopRow01;
		private int oMatBottomRow01;
		private bool AutoBatch;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="formUID"></param>
		public override void LoadForm(string formUID)
		{
			try
			{
				oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
				oForm.Freeze(true);
				SubMain.Add_Forms(this, formUID, "S41");
				oMat01 = oForm.Items.Item("35").Specific;
				oMat02 = oForm.Items.Item("3").Specific;
				oMatTopRow01 = 1;
				oMatBottomRow01 = 1;
				PS_S41_CreateItems();

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_S41_CreateItems
		/// </summary>
		private void PS_S41_CreateItems()
		{
			SAPbouiCOM.Item oNewITEM = null;

			try
			{
				oNewITEM = oForm.Items.Add("Option01", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("128").Left + 85;
				oNewITEM.Top = oForm.Items.Item("128").Top;
				oNewITEM.Height = oForm.Items.Item("128").Height;
				oNewITEM.Width = 80;

				oForm.DataSources.UserDataSources.Add("Option01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oNewITEM.Specific.DataBind.SetBound(true, "", "Option01");
				oNewITEM.Specific.Caption = "개별입고";

				oNewITEM = oForm.Items.Add("Option02", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Option01").Left + 80;
				oNewITEM.Top = oForm.Items.Item("Option01").Top;
				oNewITEM.Height = oForm.Items.Item("Option01").Height;
				oNewITEM.Width = 80;

				oForm.DataSources.UserDataSources.Add("Option02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oNewITEM.Specific.DataBind.SetBound(true, "", "Option02");
				oNewITEM.Specific.Caption = "통합입고";
				oNewITEM.Specific.GroupWith("Option01");

				oNewITEM = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Option02").Left + 80;
				oNewITEM.Top = oForm.Items.Item("Option02").Top;
				oNewITEM.Height = oForm.Items.Item("Option02").Height;
				oNewITEM.Width = 50;
				oNewITEM.Specific.Caption = "배치번호";

				oNewITEM = oForm.Items.Add("Edit01", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Static01").Left + 55;
				oNewITEM.Top = oForm.Items.Item("Static01").Top;
				oNewITEM.Height = oForm.Items.Item("Static01").Height;
				oNewITEM.Width = 50;

				oForm.DataSources.UserDataSources.Add("Edit01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oNewITEM.Specific.DataBind.SetBound(true, "", "Edit01");

				oNewITEM = oForm.Items.Add("Edit02", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Edit01").Left + 55;
				oNewITEM.Top = oForm.Items.Item("Edit01").Top;
				oNewITEM.Height = oForm.Items.Item("Edit01").Height;
				oNewITEM.Width = 50;

				oForm.DataSources.UserDataSources.Add("Edit02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oNewITEM.Specific.DataBind.SetBound(true, "", "Edit02");

				oNewITEM = oForm.Items.Add("Edit03", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Edit02").Left + 55;
				oNewITEM.Top = oForm.Items.Item("Edit02").Top;
				oNewITEM.Height = oForm.Items.Item("Edit02").Height;
				oNewITEM.Width = 50;

				oForm.DataSources.UserDataSources.Add("Edit03", SAPbouiCOM.BoDataType.dt_QUANTITY);
				oNewITEM.Specific.DataBind.SetBound(true, "", "Edit03");

				oNewITEM = oForm.Items.Add("Button01", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				oNewITEM.AffectsFormMode = false;

				oNewITEM.Left = oForm.Items.Item("Edit03").Left + 55;
				oNewITEM.Top = oForm.Items.Item("Edit03").Top - 1;
				oNewITEM.Height = oForm.Items.Item("Edit03").Height + 2;
				oNewITEM.Width = 80;

				oNewITEM.Specific.Caption = "배치번호설정";

				AutoBatch = false;

				oForm.Items.Item("Option01").Specific.Selected = true;
				oForm.Items.Item("Edit01").Enabled = false;
				oForm.Items.Item("Edit02").Enabled = false;
				oForm.Items.Item("Edit03").Enabled = false;
				oForm.Items.Item("Button01").Enabled = false;

				oMat01.Columns.Item("0").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				oMat02.Columns.Item("0").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				oForm.Items.Item("36").Visible = false;

				oNewITEM = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				oNewITEM.Top = oForm.Items.Item("2").Top;
				oNewITEM.Left = oForm.Items.Item("2").Left + 70;
				oNewITEM.Height = 12;
				oNewITEM.Width = 120;
				oNewITEM.FontSize = 12;
				oNewITEM.Specific.Caption = "Addon running";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oNewITEM);
			}
		}

		/// <summary>
		/// PS_S41_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_S41_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			int j;
			string sQry;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				for (i = 1; i <= oMat02.VisualRowCount; i++)
				{
					for (j = i + 1; j <= oMat02.VisualRowCount; j++)
					{
						if (oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim() == oMat02.Columns.Item("2").Cells.Item(j).Specific.Value.ToString().Trim())
						{
							errMessage = "동일한 배치번호가 존재합니다.";
							throw new Exception();
						}
					}

					if (string.IsNullOrEmpty(oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						continue; // for문 다음으로..
					}

					sQry = "SELECT BatchNum FROM [OIBT] WHERE ItemCode = '" + oMat01.Columns.Item("5").Cells.Item(oMatTopRow01).Specific.Value.ToString().Trim() + "' AND Quantity > 0";
					oRecordSet.DoQuery(sQry);

					for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
					{
						if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim())
						{
							errMessage = "이미 존재하는 배치번호 입니다.";
							throw new Exception();
						}
						oRecordSet.MoveNext();
					}
					//작업일보에 등록된 작업지시의 투입품, 멀티게이지의 경우만 해당된다.
					sQry = "SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry IN(SELECT U_PP030HNo FROM [@PS_PP040L] WHERE U_OrdGbn IN('104','107')) AND U_ItemCode = '" + oMat01.Columns.Item("5").Cells.Item(oMatTopRow01).Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
					{
						if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim())
						{
							//해당배치는 이미 생산에 투입된배치
							errMessage = "이미 생산에 투입된 배치번호 입니다.";
							throw new Exception();
						}
						oRecordSet.MoveNext();
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
				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			int ValidBatch;
			int StartValue;
			int EndValue;
			string BatchNum;
			string errMessage = string.Empty;
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_S41_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_S41_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
					else if (pVal.ItemUID == "Option01")
					{
						for (i = 1; i <= oMat02.VisualRowCount; i++)
						{
							oMat02.Columns.Item("3").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("1293");
						}

						oForm.Items.Item("Edit01").Specific.Value = "";
						oForm.Items.Item("Edit02").Specific.Value = "";
						oForm.Items.Item("Edit03").Specific.Value = "";
						oMat02.Columns.Item("2").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("Edit01").Enabled = false;
						oForm.Items.Item("Edit02").Enabled = false;
						oForm.Items.Item("Edit03").Enabled = false;
						oForm.Items.Item("Button01").Enabled = false;

						AutoBatch = false;
					}
					else if (pVal.ItemUID == "Option02")
					{
						for (i = 1; i <= oMat02.VisualRowCount; i++)
						{
							oMat02.Columns.Item("3").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("1293");
						}

						oForm.Items.Item("Edit01").Specific.Value = "";
						oForm.Items.Item("Edit02").Specific.Value = "";
						oForm.Items.Item("Edit03").Specific.Value = "";
						oMat02.Columns.Item("2").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("Edit01").Enabled = true;
						oForm.Items.Item("Edit02").Enabled = true;
						oForm.Items.Item("Edit03").Enabled = true;
						oForm.Items.Item("Button01").Enabled = true;
					}
					else if (pVal.ItemUID == "Button01")
					{
						StartValue = Convert.ToInt32(codeHelpClass.Right(oForm.Items.Item("Edit01").Specific.Value.ToString().Trim(), 3));
						EndValue = Convert.ToInt32(codeHelpClass.Right(oForm.Items.Item("Edit02").Specific.Value.ToString().Trim(), 3));

						if (StartValue == 0 && EndValue == 0)
						{
							errMessage = "배치번호의 범위가 올바르지 않습니다.";
							throw new Exception();
						}

						ValidBatch = EndValue - StartValue; //앞에문서 - 뒤에문서

						if (ValidBatch < 0)
						{
							errMessage = "배치번호의 범위가 올바르지 않습니다.";
							throw new Exception();
						}

						ValidBatch = (ValidBatch / 10) + 1; //10단위로 몇개생성가능한지 계산

						for (i = 1; i <= oMat02.VisualRowCount; i++)
						{
							oMat02.Columns.Item("3").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("1293");
						}

						for (i = 1; i <= ValidBatch; i++)
						{
							BatchNum = Convert.ToString(StartValue + (10 * (i - 1)));
							if (Convert.ToInt32(BatchNum) < 100)
							{
								BatchNum = "0" + BatchNum;
							}

							oMat02.Columns.Item("2").Cells.Item(i).Specific.Value = oForm.Items.Item("Edit01").Specific.Value.ToString().Trim().Substring(0, oForm.Items.Item("Edit01").Specific.Value.ToString().Trim().Length - 3) + BatchNum;

							if (ValidBatch != i)
							{
								if (string.IsNullOrEmpty(oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim()))
								{
								}
								else
								{
									if (Convert.ToDouble(oMat01.Columns.Item("39").Cells.Item(oMatTopRow01).Specific.Value.ToString().Trim()) <= Convert.ToDouble(oForm.Items.Item("Edit03").Specific.Value.ToString().Trim()) * i)
									{
										break; // TODO: might not be correct. Was : Exit For
									}
									else
									{
										oMat02.Columns.Item("5").Cells.Item(i).Specific.Value = oForm.Items.Item("Edit03").Specific.Value.ToString().Trim();
									}
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
			}
		}

		/// <summary>
		/// CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			int j;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "35")
					{
						if (pVal.Row > 0)
						{
							for (i = 1; i <= oMat02.VisualRowCount; i++)
							{
								//동일한 배치번호가 존재하는지 검사
								for (j = i + 1; j <= oMat02.VisualRowCount; j++)
								{
									if (oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim() == oMat02.Columns.Item("2").Cells.Item(j).Specific.Value.ToString().Trim())
									{
										errMessage = "동일한 배치번호가 존재합니다.";
										throw new Exception();
									}
								}
								//배치번호를 입력하지 않은경우 넘어감
								if (string.IsNullOrEmpty(oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim()))
								{
									continue;
								}

								sQry = "SELECT BatchNum FROM [OIBT] WHERE ItemCode = '" + oMat01.Columns.Item("5").Cells.Item(oMatTopRow01).Specific.Value.ToString().Trim() + "' AND Quantity > 0";
								oRecordSet.DoQuery(sQry);

								for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
								{
									if (oRecordSet.Fields.Item(0).Value == oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim())
									{
										errMessage = "이미 존재하는 배치번호 입니다.";
										throw new Exception();
									}
									oRecordSet.MoveNext();
								}
								//작업일보에 등록된 작업지시의 투입품, 멀티게이지,엔드베어링의 경우만 해당된다.
								sQry = "SELECT U_BatchNum FROM [@PS_PP030L] WHERE DocEntry IN(SELECT U_PP030HNo FROM [@PS_PP040L] WHERE U_OrdGbn IN('104','107')) AND U_ItemCode = '" + oMat01.Columns.Item("5").Cells.Item(oMatTopRow01).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
								{
									if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == oMat02.Columns.Item("2").Cells.Item(i).Specific.Value.ToString().Trim())
									{
										errMessage = "이미 생산에 투입된 배치번호 입니다.";
										throw new Exception();
									}
									oRecordSet.MoveNext();
								}
							}
							oMatTopRow01 = pVal.Row;
						}
					}
					if (pVal.ItemUID == "3")
					{
						if (pVal.Row > 0)
						{
							oMatBottomRow01 = pVal.Row;
						}
					}
				}
				else if (pVal.Before_Action == false)
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
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (oMat02.VisualRowCount <= 1)
					{
						PSH_Globals.SBO_Application.MessageBox("행을 삭제 할수 없습니다.");
						BubbleEvent = false;
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
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}

				if (pVal.ItemUID == "35")
				{
					if (pVal.Row > 0)
					{
						oMatTopRow01 = pVal.Row;
					}
				}
				if (pVal.ItemUID == "3")
				{
					if (pVal.Row > 0)
					{
						oMatBottomRow01 = pVal.Row;
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
