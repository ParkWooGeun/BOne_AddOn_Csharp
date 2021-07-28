using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산부하현황
	/// </summary>
	internal class PS_PP905 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_PP905L;  //라인(공정별부하현황)
		private SAPbouiCOM.DBDataSource oDS_PS_PP905M;  //라인(세부내역(품목별))
		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP905.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP905_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP905");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP905_CreateItems();
				PS_PP905_ResizeForm();
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
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
			}
		}

		/// <summary>
		/// PS_PP905_CreateItems
		/// </summary>
		private void PS_PP905_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_PP905L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP905M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				//품목대분류
				oForm.DataSources.UserDataSources.Add("DWkTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("DWkTime").Specific.DataBind.SetBound(true, "", "DWkTime");
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
		/// PS_PP905_ResizeForm
		/// </summary>
		private void PS_PP905_ResizeForm()
		{
			try
			{
				oForm.Items.Item("Rec01").Height = oForm.Items.Item("Mat01").Height + 20;
				oForm.Items.Item("Rec01").Width = oForm.Items.Item("Mat01").Width + 20;

				oMat01.AutoResizeColumns();
				oMat02.AutoResizeColumns();
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
		/// PS_PP905_AddMatrixRow1
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP905_AddMatrixRow1(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)
				{
					oDS_PS_PP905L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_PP905L.Offset = oRow;
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
		/// PS_PP905_AddMatrixRow2
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP905_AddMatrixRow2(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)
				{
					oDS_PS_PP905M.InsertRecord(oRow);
				}
				oMat02.AddRow();
				oDS_PS_PP905M.Offset = oRow;
				oMat02.LoadFromDataSource();
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
		/// PS_PP905_MTX01
		/// </summary>
		private void PS_PP905_MTX01()
		{
			int ErrNum = 0;
			int loopCount;
			string sQry;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				oForm.Freeze(true);

				sQry = "EXEC PS_PP905_01";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP905L.InsertRecord(loopCount);
					}
					oDS_PS_PP905L.Offset = loopCount;

					oDS_PS_PP905L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
					oDS_PS_PP905L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim()); //공정코드
					oDS_PS_PP905L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("CpName").Value.ToString().Trim()); //공정명
					oDS_PS_PP905L.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("M_Time").Value.ToString().Trim()); //장비
					oDS_PS_PP905L.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("T_Time").Value.ToString().Trim()); //공구
					oDS_PS_PP905L.SetValue("U_ColQty03", loopCount, oRecordSet.Fields.Item("P_Time").Value.ToString().Trim()); //부품
					oDS_PS_PP905L.SetValue("U_ColQty04", loopCount, oRecordSet.Fields.Item("G_Time").Value.ToString().Trim()); //게이지
					oDS_PS_PP905L.SetValue("U_ColQty05", loopCount, oRecordSet.Fields.Item("J_Time").Value.ToString().Trim()); //몰드
					oDS_PS_PP905L.SetValue("U_ColQty06", loopCount, oRecordSet.Fields.Item("LoadTime").Value.ToString().Trim()); //부하시간

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP905_MTX02
		/// </summary>
		/// <param name="prmCpCode"></param>
		/// <param name="prmItemClass"></param>
		private void PS_PP905_MTX02(string prmCpCode, string prmItemClass)
		{
			int ErrNum = 0;
			int loopCount;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				oForm.Freeze(true);

				sQry = "EXEC PS_PP905_02 '" + prmCpCode + "','" + prmItemClass + "'";
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat02.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP905M.InsertRecord(loopCount);
					}

					oDS_PS_PP905M.Offset = loopCount;
					oDS_PS_PP905M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
					oDS_PS_PP905M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //작번
					oDS_PS_PP905M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("SubCode").Value.ToString().Trim()); //서브작번
					oDS_PS_PP905M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_PP905M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("Spec").Value.ToString().Trim()); //규격
					oDS_PS_PP905M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("Unit").Value.ToString().Trim()); //단위
					oDS_PS_PP905M.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("Qty").Value.ToString().Trim()); //수량
					oDS_PS_PP905M.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("StdTime").Value.ToString().Trim()); //표준공수
					oDS_PS_PP905M.SetValue("U_ColQty04", loopCount, oRecordSet.Fields.Item("SWTime").Value.ToString().Trim()); //공수
					oDS_PS_PP905M.SetValue("U_ColQty03", loopCount, oRecordSet.Fields.Item("WkTime").Value.ToString().Trim()); //실동공수
					oDS_PS_PP905M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("WkDate").Value.ToString().Trim()); //최종작업일자

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("결과가 존재하지 않습니다.", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
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
			switch (pVal.EventType) {
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:				//1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:					//2
					//Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:				//5
					//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK:					    //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:				//7
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:		//8
					//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE:					//10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:				//11
					//Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:				//18
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:			//19
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:				//20
					Raise_EVENT_RESIZE(FormUID,  pVal, BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:			//27
					//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:					//3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:					//4
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:				//17
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "btnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							oMat02.Clear();          //실동공수 정보 초기화
							oDS_PS_PP905M.Clear();

							PS_PP905_MTX01();		 //매트릭스에 데이터 로드
						}
					}
					else if (pVal.ItemUID == "Link04")
					{
						PS_PP030 oTempClass = new PS_PP030();
						oTempClass.LoadForm(oForm.Items.Item("WODocNum").Specific.VALUE.ToString().Trim());
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					if (pVal.ItemUID == "Folder01")     //Folder01이 선택되었을 때
					{
						oForm.PaneLevel = 1;
					}
					else if (pVal.ItemUID == "Folder02")     //Folder02가 선택되었을 때
					{
						oForm.PaneLevel = 2;
					}
					else if (pVal.ItemUID == "Folder03")    //Folder03이 선택되었을 때
					{
						oForm.PaneLevel = 3;
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
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat01.SelectRow(pVal.Row, true, false);

							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
						}
					}
					else if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row > 0)
						{
							oMat02.SelectRow(pVal.Row, true, false);

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
				else if (pVal.BeforeAction == false)
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int ErrNum = 0;
			string CpCode;
			string ItemClass;
			string ColName = String.Empty;

			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					//공정별 부하현황
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;  //정렬
							oMat01.FlushToDataSource();
						}
						else
						{
							CpCode = oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.VALUE;
							ItemClass = codeHelpClass.Left(pVal.ColUID, 1);

							//공정 열에서 더블클릭하면
							if (ItemClass == "C" || ItemClass == "L")
							{
								ErrNum = 1;
								throw new Exception();
							}
							else
							{
								PS_PP905_MTX02(CpCode, ItemClass);
							}

							oForm.Items.Item("Folder02").Specific.Select();		//작업지시정보 TAB 선택
							oForm.Items.Item("CpCode").Specific.VALUE = oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.VALUE;
							oForm.Items.Item("CpName").Specific.VALUE = oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.VALUE;

							if ( codeHelpClass.Left(pVal.ColUID, 1) == "M")
							{
								ColName = "장비";
							}
							else if (codeHelpClass.Left(pVal.ColUID, 1) == "T")
							{
								ColName = "공구";
							}
							else if (codeHelpClass.Left(pVal.ColUID, 1) == "P")
							{
								ColName = "부품";
							}
							else if (codeHelpClass.Left(pVal.ColUID, 1) == "G")
							{
								ColName = "게이지";
							}
							else if (codeHelpClass.Left(pVal.ColUID, 1) == "J")
							{
								ColName = "몰드";
							}

							oForm.Items.Item("ItemClass").Specific.VALUE = ColName;
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					dataHelpClass.MDC_GF_Message("장비, 공구, 부품, 게이지, 몰드 열에서만 더블클릭 하십시오..", "W");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
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
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP905_ResizeForm();
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
				else if (pVal.ItemUID == "Mat02")
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
			finally
			{
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP905L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP905M);
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							if (oForm.PaneLevel == 1)
							{
								PS_PP905_AddMatrixRow1(oMat01.VisualRowCount, false);
							}
							else if (oForm.PaneLevel == 2)
							{
								PS_PP905_AddMatrixRow2(oMat02.VisualRowCount, false);
							}
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);

							if (oForm.PaneLevel == 1)
							{
								oDS_PS_PP905L.RemoveRecord(oDS_PS_PP905L.Size - 1);
								oMat01.LoadFromDataSource();
							}
							else if (oForm.PaneLevel == 2)
							{
								oDS_PS_PP905M.RemoveRecord(oDS_PS_PP905M.Size - 1);
								oMat02.LoadFromDataSource();
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
	}
}
