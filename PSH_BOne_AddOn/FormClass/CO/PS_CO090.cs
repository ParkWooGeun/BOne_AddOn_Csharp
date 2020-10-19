using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통계주요지표 값 입력
	/// </summary>
	internal class PS_CO090 : PSH_BaseClass
	{
		public string oFormUniqueID01;
		public SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_CO090H;  //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO090L;  //등록라인
		private int oLastColRow01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO090.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_CO090_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO090");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";				// UDO방식일때

				oForm.EnableMenu(("1293"), true);				// 행삭제
				oForm.EnableMenu(("1287"), true);				// 복제

				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();
				AddMatrixRow(0, true);

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			int loopCount = 0;
			Double Total = 0;
			string YM4 = String.Empty;
			string ATCode = String.Empty;
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
							if (pval.ItemUID == "1")
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
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
										YM4 = oDS_PS_CO090H.GetValue("U_YM", 0).ToString().Trim().Substring(2, 4);	//오른쪽부터 2자리 자를 때 ???  
										ATCode = oDS_PS_CO090H.GetValue("U_ATCode", 0).ToString().Trim();

										oDS_PS_CO090H.SetValue("Code", 0, YM4 + ATCode);
										oDS_PS_CO090H.SetValue("Name", 0, YM4 + ATCode);
									}
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							if (pval.CharPressed == 9)
							{
								if (pval.ColUID == "ReCCCode")
								{
									if (string.IsNullOrEmpty(oMat01.Columns.Item("ReCCCode").Cells.Item(pval.Row).Specific.VALUE))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
										BubbleEvent = false;
									}
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:						//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							    //6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:						//7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:				//8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:							//10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:						//11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:						//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:					//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:						//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:					//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:							//3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: 						//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:						//17
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
							if (pval.ItemUID == "Btn01")
							{
								LoadData();
							}
							if (pval.ItemUID == "1")    // 저장 후 Matrix row 한 줄 추가
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
								{
									oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
									PSH_Globals.SBO_Application.ActivateMenuItem("1282");
								}
								else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
								{
									FormItemEnabled();
									AddMatrixRow(0, true);
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:						//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							    //6
							// 전기일 시작 종료
							sQry = "select F_RefDate,T_RefDate from OFPR Where CONVERT(CHAR(6),F_RefDate,112) = '" + oForm.Items.Item("YM").Specific.VALUE.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("FRefDate").Specific.VALUE = Convert.ToDateTime(oRecordSet.Fields.Item(0).Value.ToString().Trim()).ToString("yyyyMMdd");
							oForm.Items.Item("TRefDate").Specific.VALUE = Convert.ToDateTime(oRecordSet.Fields.Item(1).Value.ToString().Trim()).ToString("yyyyMMdd");
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:						//7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:				//8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:							//10
							//통계주요지표코드가 바뀌면 한 행을 추가
							if (pval.ItemChanged == true)
							{
								if (pval.ColUID == "ReCCCode")
								{
									FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID);
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:						//11
							//합계 계산 로직 수정(2011.09.26 송명규)
							for (loopCount = 0; loopCount <= oMat01.RowCount - 1; loopCount++)
							{
								Total = Total + Convert.ToDouble(oDS_PS_CO090L.GetValue("U_Qty", loopCount));
							}
							oForm.Items.Item("Total").Specific.VALUE = Total;
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:						//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:					//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:    					//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:					//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:							//3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: 						//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO090H); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO090L); //메모리 해제
							SubMain.Remove_Forms(oFormUniqueID01);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// Raise_FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			int i = 0;

			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							if (oMat01.RowCount != oMat01.VisualRowCount)
							{
								for (i = 1; i <= oMat01.VisualRowCount; i++)
								{
									oMat01.Columns.Item("LineId").Cells.Item(i).Specific.VALUE = i;
								}
								oMat01.FlushToDataSource();								// DBDataSource에 레코드가 한줄 더 생긴다.
								oDS_PS_CO090L.RemoveRecord(oDS_PS_CO090L.Size - 1);		// 레코드 한 줄을 지운다.
								oMat01.LoadFromDataSource();							// DBDataSource를 매트릭스에 올리고
								if (oMat01.RowCount == 0)
								{
									AddMatrixRow(0, false );
								}
								else
								{
									if (!string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_ReCCCode", oMat01.RowCount - 1).ToString().Trim()))
									{
										//                            Call AddMatrixRow(1, oMat01.RowCount, True)
										AddMatrixRow(oMat01.RowCount, false);
									}
								}
							}
							break;
						case "1281":							//찾기
							AddMatrixRow(0, true);				// UDO방식
							oForm.DataBrowser.BrowseBy = "Code";		////UDO방식일때        '찾기버튼 클릭시 Matrix에 행 추가
							break;
						case "1282":							//추가
							AddMatrixRow(0, true);				//UDO방식
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
							break;
						//복제(2012.01.02 송명규 추가)
						case "1287":
							oForm.Freeze(true);
							oDS_PS_CO090H.SetValue("Code", 0, "");

							for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
							{
								oMat01.FlushToDataSource();
								oDS_PS_CO090L.SetValue("Code", i, "");
								oMat01.LoadFromDataSource();
							}
							oForm.Freeze(false);
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
			}
		}

		/// <summary>
		/// CreateItems
		/// </summary>
		/// <returns></returns>
		private bool CreateItems()
		{
			bool functionReturnValue = false;

			try
			{
				oForm.Freeze(true);

				oDS_PS_CO090H = oForm.DataSources.DBDataSources.Item("@PS_CO090H");
				oDS_PS_CO090L = oForm.DataSources.DBDataSources.Item("@PS_CO090L");

				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM, 15);
				oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			oForm.Freeze(false);
			return functionReturnValue;
		}

		/// <summary>
		/// FormItemEnabled
		/// </summary>
		public void FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					oForm.EnableMenu("1281", true);					//찾기
					oForm.EnableMenu("1282", false);				//추가
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
				{
					oForm.EnableMenu("1281", false);				//찾기
					oForm.EnableMenu("1282", true);					//추가
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
				{
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
		/// AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		public void AddMatrixRow(int oRow, bool RowIserted = false)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false)    // 행추가여부
				{
					oRow = oMat01.RowCount;
					oDS_PS_CO090L.InsertRecord((oRow));
				}
				oMat01.AddRow();
				oDS_PS_CO090L.Offset = oRow;
				oDS_PS_CO090L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
				oMat01.LoadFromDataSource();
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
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
				if ((eventInfo.BeforeAction == true))
				{
					//작업
				}
				else if ((eventInfo.BeforeAction == false))
				{
					//작업
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
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
		{
			int i = 0;

			try
			{
				if ((oLastColRow01 > 0))
				{
					if (pval.BeforeAction == true)
					{
						// 행삭제전 행삭제가능여부검사
					}
					else if (pval.BeforeAction == false)
					{
						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
						}
						oMat01.FlushToDataSource();
						oDS_PS_CO090L.RemoveRecord(oDS_PS_CO090L.Size - 1);
						oMat01.LoadFromDataSource();
						if (oMat01.RowCount == 0)
						{
							AddMatrixRow(0);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_CECode", oMat01.RowCount - 1).ToString().Trim()))
							{
								AddMatrixRow(oMat01.RowCount);
							}
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
			}
		}

		/// <summary>
		/// MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			int i = 0;
			int ErrNum = 0;

			try
			{
				// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
				oMat01.FlushToDataSource();

				// 라인
				if (oMat01.VisualRowCount <= 1)
				{
					ErrNum = 1;
					throw new Exception();
				}

				if (oMat01.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
					{
						oDS_PS_CO090L.Offset = i;

						if (string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_ReCCCode", i).ToString().Trim()))
						{
							ErrNum = 2;
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_ReCCName", i).ToString().Trim()))
						{
							ErrNum = 3;
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_Qty", i).ToString().Trim()))
						{
							ErrNum = 4;
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_Unit", i).ToString().Trim()))
						{
							ErrNum = 5;
							throw new Exception();
						}
					}

					if (string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_ReCCCode", oMat01.VisualRowCount - 1)))
					{
						oDS_PS_CO090L.RemoveRecord(oMat01.VisualRowCount - 1);
					}
				}
				// 행을 삭제하였으니 DB데이터 소스를 다시 가져온다
				oMat01.LoadFromDataSource();
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("라인데이타가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("Receiver CC는 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("CostCenter Name은 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 4)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("총수량은 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 5)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("단위는 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				functionReturnValue = false;
			}
			return functionReturnValue;
		}

		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			int ErrNum = 0;

			try
			{
				if ( string.IsNullOrEmpty(oDS_PS_CO090H.GetValue("U_YM", 0).ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
				if ( string.IsNullOrEmpty(oDS_PS_CO090H.GetValue("U_FRefDate", 0).ToString().Trim()))
				{
					ErrNum = 2;
					throw new Exception();
				}
				if ( string.IsNullOrEmpty(oDS_PS_CO090H.GetValue("U_TRefDate", 0).ToString().Trim()))
				{
					ErrNum = 3;
					throw new Exception();
				}
				if ( string.IsNullOrEmpty(oDS_PS_CO090H.GetValue("U_ATCode", 0).ToString().Trim()))
				{
					ErrNum = 4;
					throw new Exception();
				}
				
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("마감년월을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("전기일(시작)을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("전기일(종료)을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 4)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("통계주요지표선택을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				functionReturnValue = false;
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Raise_FormDataEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="BusinessObjectInfo"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
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
			}
		}

		/// <summary>
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
		{
			string sQry = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat01.FlushToDataSource();

				switch (oCol)
				{
					case "ReCCCode":
						oForm.Freeze(true);
						oMat01.FlushToDataSource();
						oDS_PS_CO090L.Offset = oRow - 1;

						sQry = "SELECT PrcName FROM OPRC WHERE PrcCode = '" + oMat01.Columns.Item("ReCCCode").Cells.Item(oRow).Specific.VALUE.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_CO090L.SetValue("U_ReCCName", oRow - 1, oRecordSet.Fields.Item("PrcName").Value.ToString().Trim());
						
						oMat01.LoadFromDataSource();
						if (oRow == oMat01.RowCount & !string.IsNullOrEmpty(oDS_PS_CO090L.GetValue("U_ReCCCode", oRow - 1).ToString().Trim()))
						{
							// 다음 라인 추가
							AddMatrixRow(0, false);
							oMat01.Columns.Item("ReCCCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		public void ComboBox_Setting()
		{
			string sQry = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 메트릭스에 있는 콤보박스 셋팅
				sQry = "SELECT U_Minor,U_CdName From [@PS_SY001L] WHERE Code = 'C001' Order by U_Minor";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oMat01.Columns.Item("Unit").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				sQry = "select U_ATCode, U_ATName from [@PS_CO060L] Order by U_ATCode";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("ATCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
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
		/// LoadData
		/// </summary>
		public void LoadData()
		{
			int i = 0;
			double TotalAmount = 0;
			string iFrDate = string.Empty;
			string iToDate = string.Empty;

			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				sQry = "SELECT U_DataProc FROM [@PS_CO060L] WHERE U_ATCode = '" + oForm.Items.Item("ATCode").Specific.VALUE.ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				iFrDate = oForm.Items.Item("FRefDate").Specific.VALUE.ToString().Trim();
				iToDate = oForm.Items.Item("TRefDate").Specific.VALUE.ToString().Trim();


				sQry = "EXEC [" + oRecordSet.Fields.Item(0).Value.ToString().Trim() + "] '" + iFrDate + "','" + iToDate + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_CO090L.Clear();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_CO090L.Size)
					{
						oDS_PS_CO090L.InsertRecord((i));
					}

					oMat01.AddRow();
					oDS_PS_CO090L.Offset = i;
					oDS_PS_CO090L.SetValue("LineId", i, Convert.ToString(i + 1));
					oDS_PS_CO090L.SetValue("U_ReCCCode", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_CO090L.SetValue("U_ReCCName", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_CO090L.SetValue("U_Qty", i,  oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_CO090L.SetValue("U_Unit", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());

					TotalAmount = TotalAmount + Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim());

					oRecordSet.MoveNext();
				}
				oForm.DataSources.UserDataSources.Item("Total").Value = Convert.ToString(TotalAmount);
				oMat01.LoadFromDataSource();
				AddMatrixRow(i, false);
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
	}
}
