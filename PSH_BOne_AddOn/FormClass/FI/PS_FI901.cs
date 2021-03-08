using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// IFRS - 매입매출채무채권추출
	/// </summary>
	internal class PS_FI901 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Grid oGrid01;

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI901.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID01 = "PS_FI901_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI901");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;

				oForm.Freeze(true);
				CreateItems();
				oForm.EnableMenu("1283", false);              //// 제거
				oForm.EnableMenu("1284", false);              //// 취소
				oForm.EnableMenu("1287", false);              //// 복원
				oForm.EnableMenu("1293", false);              //// 행삭제
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
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 그리드 개체 할당
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid01.Item.Enabled = false;
				oForm.DataSources.DataTables.Add("Grid01");

				//사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "");
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				// DocType
				oForm.Items.Item("DocType").Specific.ValidValues.Add("1", "매입");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("2", "매출");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("3", "채무");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("4", "채권");

				oForm.Items.Item("Company").Specific.Value = "PSH"; //회사
				oForm.Items.Item("AcctYear").Specific.Value = DateTime.Now.ToString("yyyy"); //회계년도
				oForm.Items.Item("AcctMon").Specific.Value = Convert.ToDateTime(DateTime.Now).AddMonths(-1).ToString("MM"); //회계월(전월표기)
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
		/// DrawGrid
		/// </summary>
		private void DrawGrid()
		{
			int i = 0;
			string sColsTitle = string.Empty;
			oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
			try
			{
				for (i = 0; i <= oGrid01.Columns.Count - 1; i++)
				{
					sColsTitle = oGrid01.Columns.Item(i).TitleObject.Caption;

					if (oGrid01.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
					{
						oGrid01.Columns.Item(i).RightJustified = true;
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
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry = string.Empty;

			int sReturnValue = 0;
			string AcctMon = string.Empty;
			string Company = string.Empty;
			string DocType = string.Empty;
			string AcctYear = string.Empty;
			string BPLId = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					switch (pVal.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
							if (pVal.ItemUID == "Btn01" || pVal.ItemUID == "Btn02" || pVal.ItemUID == "Btn03")
							{
								Company = oForm.Items.Item("Company").Specific.Value.ToString().Trim();
								DocType = oForm.Items.Item("DocType").Specific.Selected.Value.ToString().Trim();
								AcctYear = oForm.Items.Item("AcctYear").Specific.Value.ToString().Trim();
								AcctMon = oForm.Items.Item("AcctMon").Specific.Value.ToString().Trim();
								BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();

								if (string.IsNullOrEmpty(BPLId))
								{
									BPLId = "%";
								}

								if (pVal.ItemUID == "Btn01")
								{
									sQry = " Select * From [ZFI020] Where ";
									sQry += "Company = '" + Company + "' And ";
									sQry += "DocType = '" + DocType + "' And ";
									sQry += "AcctYear = '" + AcctYear + "' And ";
									sQry += "AcctMon = '" + AcctMon + "'";
									oRecordSet.DoQuery(sQry);

									if (oRecordSet.RecordCount > 0)
									{
										sReturnValue = PSH_Globals.SBO_Application.MessageBox("해당 조건의 데이터가 존재합니다. 바꾸시겠습니까?", 1, "&확인", "&취소");

									}
									else
									{
										sReturnValue = PSH_Globals.SBO_Application.MessageBox("해당 조건의 데이터를 저장하시겠습니까?", 1, "&확인", "&취소");
									}

									switch (sReturnValue)
									{
										case 1:
											if (oRecordSet.RecordCount > 0)
											{
												sQry = " Delete [ZFI020] Where ";
												sQry += "Company = '" + Company + "' And ";
												sQry += "DocType = '" + DocType + "' And ";
												sQry += "AcctYear = '" + AcctYear + "' And ";
												sQry += "AcctMon = '" + AcctMon + "'";
												oRecordSet.DoQuery(sQry);
											}
											sQry = "EXEC PS_FI901_01 '" + Company + "', '" + AcctYear + "', '" + AcctMon + "', '" + DocType + "'";
											oRecordSet.DoQuery(sQry);
											PSH_Globals.SBO_Application.StatusBar.SetText("해당 조건의 데이터가 성공적으로 저장되었습니다. 데이터를 확인해보세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

											sQry = "EXEC PS_FI901_02 '" + Company + "', '" + AcctYear + "', '" + AcctMon + "', '" + DocType + "', '" + BPLId + "'";
											oForm.DataSources.DataTables.Item(0).ExecuteQuery(sQry);
											oGrid01.DataTable = oForm.DataSources.DataTables.Item("Grid01");

											DrawGrid();
											break;
										case 2:
											PSH_Globals.SBO_Application.MessageBox("실행이 취소되었습니다.");
											BubbleEvent = false;
											break;
									}
								}
								else if (pVal.ItemUID == "Btn02")
								{
									sQry = "EXEC PS_FI901_02 '" + Company + "', '" + AcctYear + "', '" + AcctMon + "', '" + DocType + "', '" + BPLId + "'";
									oForm.DataSources.DataTables.Item(0).ExecuteQuery(sQry);
									oGrid01.DataTable = oForm.DataSources.DataTables.Item("Grid01");

									DrawGrid();
								}
								else if (pVal.ItemUID == "Btn03")
								{
									sQry = " Select * From [ZFI020] Where ";
									sQry += "Company = '" + Company + "' And ";
									sQry += "DocType = '" + DocType + "' And ";
									sQry += "AcctYear = '" + AcctYear + "' And ";
									sQry += "AcctMon = '" + AcctMon + "'";
									oRecordSet.DoQuery(sQry);

									if (oRecordSet.RecordCount > 0)
									{
										sReturnValue = PSH_Globals.SBO_Application.MessageBox("해당 조건의 데이터가 존재합니다. 삭제하시겠습니까?", 1, "&확인", "&취소");
										switch (sReturnValue)
										{
											case 1:
												sQry = " Delete [ZFI020] Where ";
												sQry += "Company = '" + Company + "' And ";
												sQry += "DocType = '" + DocType + "' And ";
												sQry += "AcctYear = '" + AcctYear + "' And ";
												sQry += "AcctMon = '" + AcctMon + "'";
												oRecordSet.DoQuery(sQry);
												PSH_Globals.SBO_Application.StatusBar.SetText("해당 조건의 데이터가 성공적으로 삭제되었습니다. 데이터를 확인해보세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

												sQry = "EXEC PS_FI901_02 '" + Company + "', '" + AcctYear + "', '" + AcctMon + "', '" + DocType + "', '" + BPLId + "'";
												oForm.DataSources.DataTables.Item(0).ExecuteQuery(sQry);
												oGrid01.DataTable = oForm.DataSources.DataTables.Item("Grid01");

												DrawGrid();
												break;
											case 2:
												PSH_Globals.SBO_Application.MessageBox("실행이 취소되었습니다.");
												BubbleEvent = false;
												break;
										}
									}
									else
									{
										PSH_Globals.SBO_Application.MessageBox("해당 조건의 데이터가 존재하지 않습니다. 데이터를 확인해보세요.");
									}
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                          //3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                         //4
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                       //5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:                              //6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                       //7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                //8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:                           //10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                        //11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                      //18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    //19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                        //20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                   //27
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                       //1
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                           //2
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                          //3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                         //4
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                       //5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:                              //6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                       //7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                //8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:                           //10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                        //11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17		
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01); //메모리 해제
							SubMain.Remove_Forms(oFormUniqueID01);
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                      //18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    //19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                        //20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                   //27
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
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
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
