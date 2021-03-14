using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// AR송장 만기일 조회 및 승인
	/// </summary>
	internal class PS_SD013 : PSH_BaseClass
	{
		public string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat;			
		private SAPbouiCOM.DBDataSource oDS_PS_SD013H;  //등록헤더		
		private SAPbouiCOM.DBDataSource oDS_PS_SD013L;	//등록라인
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oSeq;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD013.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD013_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD013");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD013_CreateItems();
				PS_SD013_ComboBox_Setting();
				PS_SD013_Initial_Setting();
				PS_SD013_FormItemEnabled();
				PS_SD013_FormClear();
				PS_SD013_AddMatrixRow(0, oMat.RowCount, true);

				oForm.EnableMenu(("1283"), false);          //삭제
				oForm.EnableMenu(("1286"), false);          //닫기
				oForm.EnableMenu(("1287"), true);           //복제
				oForm.EnableMenu(("1284"), true);           //취소
				oForm.EnableMenu(("1293"), true);           //행삭제
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
		/// PS_SD013_CreateItems
		/// </summary>
		/// <returns></returns>
		private bool PS_SD013_CreateItems()
		{
			bool functionReturnValue = false;

			try
			{
				oDS_PS_SD013H = oForm.DataSources.DBDataSources.Item("@PS_SD013H");
				oDS_PS_SD013L = oForm.DataSources.DBDataSources.Item("@PS_SD013L");
				oMat = oForm.Items.Item("Mat01").Specific;
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_SD013_ComboBox_Setting
		/// </summary>
		private void PS_SD013_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				// 사업장 리스트
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_SD013_Initial_Setting
		/// </summary>
		private void PS_SD013_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
					//Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
					Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
					//Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
					//Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
					Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
							if (PS_SD013_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_SD013_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						//대상 자료 조회
					}
					else if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD013_CheckDataDuplication() == false)
							{
								PSH_Globals.SBO_Application.MessageBox("해당월의 자료가 이미 존재합니다.");
								BubbleEvent = false;
								return;
							}
							else
							{
								PS_SD013_MTX01();
								PS_SD013_AddMatrixRow(1, oMat.RowCount, true);
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//저장 후 추가 가능처리
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_SD013_FormItemEnabled();
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
								PSH_Globals.SBO_Application.ActivateMenuItem("1291");
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_SD013_FormItemEnabled();
								PS_SD013_AddMatrixRow(1, oMat.RowCount, true);
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
		/// CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
					if(pVal.Row > 0) {
						oLast_Item_UID = pVal.ItemUID;
						oLast_Col_UID = pVal.ColUID;
						oLast_Col_Row = pVal.Row;

						oMat.SelectRow(pVal.Row, true, false);
					} else
					{
						oLast_Item_UID = pVal.ItemUID;
						oLast_Col_UID = "";
						oLast_Col_Row = 0;
					}
				}
				else if (pVal.Before_Action == false)
				{
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
		/// GOT_FOCUS 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
					oLast_Item_UID = pVal.ItemUID;
				}
				else if (pVal.Before_Action == false)
				{
					oLast_Item_UID = pVal.ItemUID;
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
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.Action_Success == true)
					{
						oSeq = 1;
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
		/// FORM_ACTIVATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_ACTIVATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (oSeq == 1)
					{
						oSeq = 0;
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
		/// FORM_RESIZE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					oMat.AutoResizeColumns();
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD013H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD013L);
				}
				else if (pVal.Before_Action == false)
				{
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
		/// FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
    	{
			try
			{
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행닫기
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
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1281":							//찾기
							PS_SD013_FormItemEnabled();
							break;
						case "1282":							//추가
							PS_SD013_FormItemEnabled();
							PS_SD013_FormClear();
							PS_SD013_AddMatrixRow(0, oMat.RowCount, true);
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							oForm.Freeze(true);
							PS_SD013_FormItemEnabled();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("AREntry").Cells.Item(oMat.VisualRowCount).Specific.VALUE))
								{
									PS_SD013_AddMatrixRow(1, oMat.RowCount, true);
								}
							}
							oMat.AutoResizeColumns();
							oForm.Freeze(false);
							break;
						case "1293":							//행닫기
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
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
		/// PS_SD013_FormItemEnabled
		/// </summary>
		private void PS_SD013_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("BPLID").Enabled = true;
					oForm.Items.Item("FrDt").Enabled = true;
					oForm.Items.Item("ToDt").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("BPLID").Enabled = true;
					oForm.Items.Item("FrDt").Enabled = true;
					oForm.Items.Item("ToDt").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("BPLID").Enabled = false;
					oForm.Items.Item("FrDt").Enabled = false;
					oForm.Items.Item("ToDt").Enabled = false;
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
		/// PS_SD013_AddMatrixRow
		/// </summary>
		/// <param name="oSeq"></param>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD013_AddMatrixRow(short oSeq, int oRow, bool RowIserted = false)
		{
			try
			{
				switch (oSeq)
				{
					case 0:
						oMat.AddRow();
						// 매트릭스에 새로운 Row를 추가한다.
						oDS_PS_SD013L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
					case 1:
						oDS_PS_SD013L.InsertRecord(oRow);
						oDS_PS_SD013L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
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
		/// PS_SD013_FormClear
		/// </summary>
		private void PS_SD013_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD013'", "");

				if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
				{
					oDS_PS_SD013H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_SD013H.SetValue("DocEntry", 0, DocEntry);
				}
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oDS_PS_SD013H.SetValue("U_FrDt", 0, Convert.ToString(DateTime.Today.AddMonths(-1).ToString("yyyyMM01")));   //전월 1일
				oDS_PS_SD013H.SetValue("U_ToDt", 0, Convert.ToString(DateTime.Today.AddMonths(0).AddDays(0 - DateTime.Today.Day).ToString("yyyyMMdd")));   //전월 말일
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
		/// PS_SD013_MTX01
		/// </summary>
		private void PS_SD013_MTX01()
		{
			int i;
			int ErrNum = 0;
			string sQry;
			string BPLId;
			string FrDt;
			string ToDt;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				BPLId = oForm.Items.Item("BPLID").Specific.Selected.VALUE.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.VALUE.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.VALUE.ToString().Trim();

				sQry = " EXEC [PS_SD013_01] '";
				sQry += BPLId + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += PSH_Globals.SBO_Application.Company.UserName + "'";     //dev01, PSH45 등의 UserName
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_SD013L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}

				oForm.Freeze(true);

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_SD013L.Size)
					{
						oDS_PS_SD013L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_SD013L.Offset = i;
					oDS_PS_SD013L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_SD013L.SetValue("U_AREntry", i, oRecordSet.Fields.Item("AREntry").Value.ToString().Trim());              //AR송장번호
					oDS_PS_SD013L.SetValue("U_CardName", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());            //거래처
					oDS_PS_SD013L.SetValue("U_ItemCode", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());            //품목코드
					oDS_PS_SD013L.SetValue("U_ItemName", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());            //품목명
					oDS_PS_SD013L.SetValue("U_Quantity", i, oRecordSet.Fields.Item("Quantity").Value.ToString().Trim());            //수량
					oDS_PS_SD013L.SetValue("U_Price", i, oRecordSet.Fields.Item("Price").Value.ToString().Trim());                  //단가
					oDS_PS_SD013L.SetValue("U_Amount", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());                //금액
					oDS_PS_SD013L.SetValue("U_VatAmt", i, oRecordSet.Fields.Item("VatAmt").Value.ToString().Trim());                //세액
					oDS_PS_SD013L.SetValue("U_DueDate", i, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //만기일
					oDS_PS_SD013L.SetValue("U_UserCode", i, oRecordSet.Fields.Item("UserCode").Value.ToString().Trim());            //등록자ID
					oDS_PS_SD013L.SetValue("U_UserName", i, oRecordSet.Fields.Item("UserName").Value.ToString().Trim());            //등록자성명

					oRecordSet.MoveNext();

					ProgBar01.Value = ProgBar01.Value + 1;
					ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();

			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				ProgBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD013_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_SD013_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			short ErrNum = 0;

			try
			{
				if ( string.IsNullOrEmpty(oDS_PS_SD013H.GetValue("U_BPLID", 0)) || string.IsNullOrEmpty(oDS_PS_SD013H.GetValue("U_FrDt", 0)) || string.IsNullOrEmpty(oDS_PS_SD013H.GetValue("U_ToDt", 0)))
				{
					ErrNum = 1;
					throw new Exception();
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("사업장, 조회일자는 필수입력 사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_SD013_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_SD013_MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			short ErrNum = 0;

			try
			{
				oMat.FlushToDataSource();
				//라인
				if (oMat.VisualRowCount <= 1)
				{
					ErrNum = 1;
					throw new Exception();
				}

				oDS_PS_SD013L.RemoveRecord(oDS_PS_SD013L.Size - 1);
				oMat.LoadFromDataSource();
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_SD013_CheckDataDuplication
		/// </summary>
		/// <returns></returns>
		private bool PS_SD013_CheckDataDuplication()
		{
			bool functionReturnValue = false;
			string sQry;
			string BPLId;
			string Creator;
			string FrDt;
			string ToDt;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLId = oForm.Items.Item("BPLID").Specific.Selected.VALUE.ToString().Trim();
				Creator = PSH_Globals.SBO_Application.Company.UserName.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.VALUE.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.VALUE.ToString().Trim();

				sQry = "  SELECT COUNT(*) AS [Cnt]";
				sQry += " FROM  [@PS_SD013H] AS T0";
				sQry += " WHERE T0.U_BPLID = '" + BPLId + "'";
				sQry += "   AND T0.Creator = '" + Creator + "'";
				sQry += "   AND T0.U_FrDt = '" + FrDt + "'";
				sQry += "   AND T0.U_ToDt = '" + ToDt + "'";
				sQry += "   AND T0.Status = 'O'";

				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item("Cnt").Value >= 1)
				{
					functionReturnValue = false;
				}
				else
				{
					functionReturnValue = true;
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
			return functionReturnValue;
		}
	}
}
