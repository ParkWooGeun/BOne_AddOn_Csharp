using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 코스트센터비용집계
	/// </summary>
	internal class PS_CO080 : PSH_BaseClass
	{
		public string oFormUniqueID01;
		public SAPbouiCOM.Matrix oMat01;
			
		private SAPbouiCOM.DBDataSource oDS_PS_CO080H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO080L; //등록라인
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
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO080.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_CO080_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_CO080");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";                // UDO방식일때

				oForm.EnableMenu(("1293"), true);               // 행삭제
				oForm.EnableMenu(("1287"), true);               // 복제

				oForm.Freeze(true);
				CreateItems();
				AddMatrixRow(0, true);
				ComboBox_Setting();
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
			string Code = String.Empty;
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
							if (pval.ItemUID == "1")
							{
								Code = oDS_PS_CO080H.GetValue("U_YM", 0).ToString().Trim() + oDS_PS_CO080H.GetValue("U_BPLId", 0).ToString().Trim();
								oDS_PS_CO080H.SetValue("Code", 0, Code);
								oDS_PS_CO080H.SetValue("Name", 0, Code);
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
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:      					//6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:   		    		//20
							break; 
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:   					//3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:					//17
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
							if (pval.ItemUID == "Btn01")
							{
								LoadData();
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							//6
							// 전기일 시작 종료
							sQry = "select F_RefDate,T_RefDate from OFPR Where CONVERT(CHAR(6),F_RefDate,112) = '" + oForm.Items.Item("YM").Specific.VALUE.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("FRefDate").Specific.VALUE = Convert.ToDateTime(oRecordSet.Fields.Item(0).Value.ToString().Trim()).ToString("yyyyMMdd");
							oForm.Items.Item("TRefDate").Specific.VALUE = Convert.ToDateTime(oRecordSet.Fields.Item(1).Value.ToString().Trim()).ToString("yyyyMMdd");
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
							//코스트센터가 바뀌면 한 행을 추가
							if (pval.ItemChanged == true)
							{
								if (pval.ColUID == "ActCode")
								{
									FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID);
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
							AddMatrixRow(pval.Row, false);
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:    				//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:   					//3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO080H); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO080L); //메모리 해제
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
							oForm.DataBrowser.BrowseBy = "Code";							//UDO방식일때
							break;
						case "1282":							//추가
							oForm.DataBrowser.BrowseBy = "Code";							//UDO방식일때
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
								oDS_PS_CO080L.RemoveRecord(oDS_PS_CO080L.Size - 1);		// 레코드 한 줄을 지운다.
								oMat01.LoadFromDataSource();							// DBDataSource를 매트릭스에 올리고
								if (oMat01.RowCount == 0)
								{
									AddMatrixRow(0, false);
								}
								else
								{
									if (!string.IsNullOrEmpty(oDS_PS_CO080L.GetValue("U_ActCode", oMat01.RowCount - 1).ToString().Trim()))
									{
										AddMatrixRow(oMat01.RowCount, false);
									}
								}
							}
							break;
						case "1281":							//찾기
							AddMatrixRow(0, true);							//UDO방식
							oForm.DataBrowser.BrowseBy = "Code";						//UDO방식일때        '찾기버튼 클릭시 Matrix에 행 추가
							break;
						case "1282":							//추가
							//                 Call FormItemEnabled '//UDO방식
							AddMatrixRow(0, true);					 //UDO방식
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
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

				oDS_PS_CO080H = oForm.DataSources.DBDataSources.Item("@PS_CO080H");
				oDS_PS_CO080L = oForm.DataSources.DBDataSources.Item("@PS_CO080L");

				oMat01 = oForm.Items.Item("Mat01").Specific;

				oForm.DataSources.UserDataSources.Add("Amount", SAPbouiCOM.BoDataType.dt_SUM, 15);
				oForm.Items.Item("Amount").Specific.DataBind.SetBound(true, "", "Amount");
				oMat01.AutoResizeColumns();

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			oForm.Freeze(false);
			return functionReturnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		public void FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					oForm.EnableMenu("1281", true);                 //찾기
					oForm.EnableMenu("1282", false);                //추가
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
				{
					oForm.EnableMenu("1281", false);                //찾기
					oForm.EnableMenu("1282", true);                 //추가
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
				////행추가여부
				if (RowIserted == false)
				{
					oRow = oMat01.RowCount;
					oDS_PS_CO080L.InsertRecord((oRow));
				}
				oMat01.AddRow();
				oDS_PS_CO080L.Offset = oRow;
				oDS_PS_CO080L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
						oDS_PS_CO080L.RemoveRecord(oDS_PS_CO080L.Size - 1);
						oMat01.LoadFromDataSource();
						if (oMat01.RowCount == 0)
						{
							AddMatrixRow(0);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_CO080L.GetValue("U_CECode", oMat01.RowCount - 1).ToString().Trim()))
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

			int ErrNum = 0;

			try
			{
				oMat01.FlushToDataSource();
				// 라인
				if (oMat01.VisualRowCount < 1)
				{
					ErrNum = 1;
					throw new Exception();
				}

				// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
				// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
				if (oMat01.VisualRowCount > 0)
				{
					if (string.IsNullOrEmpty(oDS_PS_CO080L.GetValue("U_CoCtCode", oMat01.VisualRowCount - 1)))
					{
						oDS_PS_CO080L.RemoveRecord(oMat01.VisualRowCount - 1);
					}
				}

				//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
				oMat01.LoadFromDataSource();
				functionReturnValue = true;
			}

			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("라인데이타가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			int ErrNum = 0;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_CO080H.GetValue("Code", 0).ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_CO080H.GetValue("U_FRefDate", 0).ToString().Trim()))
				{
					ErrNum = 2;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_CO080H.GetValue("U_TRefDate", 0).ToString().Trim()))
				{
					ErrNum = 3;
					throw new Exception();
				}
				PSH_Globals.SBO_Application.MessageBox("정상등록 되었습니다.");
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("Code(Key)를 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("전기일(시작)을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("전기일(종료)을 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
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
			try
			{
				oMat01.FlushToDataSource();

				switch (oCol)
				{
					case "ActCode":
						oMat01.FlushToDataSource();
						oDS_PS_CO080L.Offset = oRow - 1;

						oForm.Freeze(true);
						oMat01.LoadFromDataSource();

						//--------------------------------------------------------------------------------------------
						if (oRow == oMat01.RowCount & !string.IsNullOrEmpty(oDS_PS_CO080L.GetValue("U_ActCode", oRow - 1).ToString().Trim()))
						{
							// 다음 라인 추가
							AddMatrixRow(0, false);
							oMat01.Columns.Item("ActCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		public void ComboBox_Setting()
		{
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "11", "재료비");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "12", "노무비");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "13", "경비");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "21", "매출");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "22", "매출원가");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "24", "판관비");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "25", "영업외수익");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "26", "영업외비용");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "27", "특별이익");
				dataHelpClass.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "28", "특별손실");

				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Class"), "PS_CO080", "Mat01", "Class", false);  // 미지막 false는 공백항목추가여부
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
		/// LoadData
		/// </summary>
		public void LoadData()
		{
			int i = 0;
			Double TotalAmount = 0;
			string iToDate = string.Empty;
			string iFrDate = string.Empty;
			string iBPLId = string.Empty;
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				iFrDate = oForm.Items.Item("FRefDate").Specific.VALUE.ToString().Trim();
				iToDate = oForm.Items.Item("TRefDate").Specific.VALUE.ToString().Trim();
				iBPLId  = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();

				sQry = "EXEC [PS_CO080_01] '" + iFrDate + "','" + iToDate + "','" + iBPLId + "'";

				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_CO080L.Clear();

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_CO080L.Size)
					{
						oDS_PS_CO080L.InsertRecord((i));
					}

					oMat01.AddRow();
					oDS_PS_CO080L.Offset = i;
					oDS_PS_CO080L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_CO080L.SetValue("U_CoCtCode", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_CO080L.SetValue("U_CoCtName", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_CO080L.SetValue("U_CoElCode", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_CO080L.SetValue("U_CoElName", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_CO080L.SetValue("U_Class", i,    oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_CO080L.SetValue("U_Amount", i,   oRecordSet.Fields.Item(5).Value.ToString().Trim());

					TotalAmount = TotalAmount + Convert.ToDouble(oRecordSet.Fields.Item(5).Value.ToString().Trim());

					oRecordSet.MoveNext();
				}
				oForm.DataSources.UserDataSources.Item("Amount").Value = Convert.ToString(TotalAmount);
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
