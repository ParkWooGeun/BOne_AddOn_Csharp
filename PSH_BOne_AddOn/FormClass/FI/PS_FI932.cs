using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 차입금등록
	/// </summary>
	internal class PS_FI932 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_FI932H;  //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_FI932L;  //등록라인

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI932.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI932_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI932");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				//화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();
				FormClear();
				Matrix_AddRow(1, 0, true);
				FormItemEnabled();

				oForm.EnableMenu("1283", false);				// 삭제
				oForm.EnableMenu("1286", false);				// 닫기
				oForm.EnableMenu("1287", false);				// 복제
				oForm.EnableMenu("1284", true);				// 취소
				oForm.EnableMenu("1293", true);             // 행삭제
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oDS_PS_FI932H = oForm.DataSources.DBDataSources.Item("@PS_FI932H");
				oDS_PS_FI932L = oForm.DataSources.DBDataSources.Item("@PS_FI932L");

				oMat01 = oForm.Items.Item("Mat01").Specific;
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
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// FormItemEnabled
		/// </summary>
		private void FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("ManageNo").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = false;
					oMat01.Columns.Item("DocDate").Editable = true;

				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("ManageNo").Enabled = false;
					oForm.Items.Item("DocDate").Enabled = true;
					oMat01.Columns.Item("DocDate").Editable = true;

				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("ManageNo").Enabled = false;
					oForm.Items.Item("DocDate").Enabled = false;
					oMat01.Columns.Item("DocDate").Editable = true;
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
		/// FormClear
		/// </summary>
		private void FormClear()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				string Code = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FI932'", "");
				if (Convert.ToDouble(Code) == 0)
				{
					oForm.Items.Item("Code").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("Code").Specific.Value = Code;
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
		/// Matrix_AddRow
		/// </summary>
		/// <param name="oMat"></param>
		/// <param name="oRow"></param>
		/// <param name="Insert_YN"></param>
		private void Matrix_AddRow(int oMat, int oRow, bool Insert_YN)
		{
			try
			{
				switch (oMat)
				{
					case 1:
						if (Insert_YN == false)
						{
							oRow = oMat01.RowCount;
							oDS_PS_FI932L.InsertRecord(oRow);
						}
						oDS_PS_FI932L.Offset = oRow;
						oDS_PS_FI932L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
						oDS_PS_FI932L.SetValue("U_DocDate", oRow, "");
						oDS_PS_FI932L.SetValue("U_Comments", oRow, "");
						oDS_PS_FI932L.SetValue("U_LoanAmt", oRow, "");
						oDS_PS_FI932L.SetValue("U_Interest", oRow, "");
						oDS_PS_FI932L.SetValue("U_RepayAmt", oRow, "");
						oDS_PS_FI932L.SetValue("U_Balance", oRow, "");
						oMat01.LoadFromDataSource();
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
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i = 0;
			Double Balance = 0;
			Double RBalance = 0;

			try
			{
				if (oUID == "Mat01")
				{
					oMat01.FlushToDataSource();
					switch (oCol)
					{
						case "DocDate":
							
							oDS_PS_FI932L.Offset = oRow - 1;    //oMat01.SetLineData oRow

							if (oRow == oMat01.RowCount && !string.IsNullOrEmpty(oDS_PS_FI932L.GetValue("U_DocDate", oRow - 1).ToString().Trim()))
							{
								Matrix_AddRow(1, 0, false);  // 다음 라인 추가
								oMat01.Columns.Item("DocDate").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
							break;
						case "LoanAmt":
						case "RepayAmt":
							oForm.Freeze(true);
							for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
							{
								oDS_PS_FI932L.Offset = i;
								Balance = RBalance + Convert.ToDouble(oDS_PS_FI932L.GetValue("U_LoanAmt", i)) - Convert.ToDouble(oDS_PS_FI932L.GetValue("U_RepayAmt", i));
								oDS_PS_FI932L.SetValue("U_Balance", i, Convert.ToString(Balance));
								RBalance = Balance;
							}
							
							oForm.Freeze(false);
							break;
					}
					oMat01.LoadFromDataSource();
					oMat01.AutoResizeColumns();
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			int ErrNum = 0;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_ManageNo", 0).ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					ErrNum = 2;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_SeqNo", 0).ToString().Trim()))
				{
					ErrNum = 3;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_LoanBank", 0).ToString().Trim()))
				{
					ErrNum = 4;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_LoanName", 0).ToString().Trim()))
				{
					ErrNum = 5;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("LoanAmt").Specific.Value))
				{
					ErrNum = 6;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_LoanSDat", 0).ToString().Trim()))
				{
					ErrNum = 7;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_LoanEDat", 0).ToString().Trim()))
				{
					ErrNum = 8;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("LoanIntr").Specific.Value))
				{
					ErrNum = 9;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FI932H.GetValue("U_RePayWay", 0).ToString().Trim()))
				{
					ErrNum = 10;
					throw new Exception();
				}
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("관리번호는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("일자는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("일자-순번은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 4)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("차입처는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 5)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("차입금명은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 6)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("차입금액은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 7)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("차입기간시작은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 8)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("차입기간종료는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 9)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("금리(%)는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 10)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("상환방법은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

			int i = 0;
			short ErrNum = 0;

			try
			{
				oMat01.FlushToDataSource();

				// 라인
				// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
				if (oMat01.VisualRowCount == 1)
				{
					ErrNum = 1;
					throw new Exception();
				}

				if (oMat01.VisualRowCount > 0)
				{
					// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
					for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
					{
						oDS_PS_FI932L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_FI932L.GetValue("U_DocDate", i).ToString().Trim()))
						{
							ErrNum = 2;
							oMat01.Columns.Item("DocDate").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							throw new Exception();
						}
					}
				}

				if (oMat01.VisualRowCount > 0)
				{
					oDS_PS_FI932L.RemoveRecord(oDS_PS_FI932L.Size - 1); // Mat1에 마지막라인(빈라인) 삭제
				}

				oMat01.LoadFromDataSource();

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("일자(라인)는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}

			return functionReturnValue;
		}

		/// <summary>
		/// 이동등록번호생성
		/// </summary>
		private void Make_ManageNo()
		{
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// Procedure 실행
				sQry = "EXEC PS_FI932_01 '" + oDS_PS_FI932H.GetValue("U_DocDate", 0) + "'";
				oRecordSet.DoQuery(sQry);

				oDS_PS_FI932H.SetValue("U_ManageNo", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
				oDS_PS_FI932H.SetValue("U_SeqNo", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
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
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					switch (pVal.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
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
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                          //3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                         //4
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
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:						//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:					//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:						//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:					//27
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
							if (pVal.ItemUID == "1")
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
								{
									oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
									PSH_Globals.SBO_Application.ActivateMenuItem("1282");
								}
								else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
								{
									FormItemEnabled();
									Matrix_AddRow(1, oMat01.RowCount, false);				//oMat01
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                           //2
                            break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                          //3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                         //4
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
							if (pVal.ItemChanged == true)
							{
								// 헤더
								if (pVal.ItemUID == "DocDate")
								{
									Make_ManageNo();
								}
								//라인
								if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "DocDate" || pVal.ColUID == "LoanAmt" || pVal.ColUID == "RepayAmt"))
								{
									FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:						//11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI932H);
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI932L);
							SubMain.Remove_Forms(oFormUniqueID01);
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:						//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:					//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:						//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:					//27
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
		/// Raise_FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i = 0;

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
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1285":							//복원
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
						case "1285":							//복원
							break;
						case "1293":							//행삭제
							if (oMat01.RowCount != oMat01.VisualRowCount)
							{
								for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
								{
									oMat01.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = i + 1;
								}
								oMat01.FlushToDataSource();
								oDS_PS_FI932L.RemoveRecord(oDS_PS_FI932L.Size - 1);	// Mat1에 마지막라인(빈라인) 삭제
								oMat01.Clear();
								oMat01.LoadFromDataSource();
							}
							break;
						case "1281":							//찾기
							FormItemEnabled();
							oForm.Items.Item("ManageNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282":							//추가
							FormItemEnabled();
							FormClear();
							Matrix_AddRow(1, 0, true);  		//oMat01
							oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							FormItemEnabled();
							if (oMat01.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat01.Columns.Item("DocDate").Cells.Item(oMat01.VisualRowCount).Specific.Value))
								{
									Matrix_AddRow(1, oMat01.RowCount, false);
								}
							}
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
		/// Raise_FormDataEvent
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
				if (eventInfo.BeforeAction == true)
				{
				}
				else if (eventInfo.BeforeAction == false)
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
	}
}
