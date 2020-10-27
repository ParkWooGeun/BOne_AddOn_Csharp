using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 분개전표 연결발행  PS_FI420
	/// </summary>
	internal class PS_FI420 : PSH_BaseClass
	{
		public string oFormUniqueID01;
		public SAPbouiCOM.Matrix oMat01;
			
		private SAPbouiCOM.DBDataSource oDS_PS_FI420L;  //등록헤더

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI420.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI420_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI420");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();
				Initialization();
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
			int i = 0;
			string Check = String.Empty;

			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
							if (pval.ItemUID == "Btn01")
							{
							}
							else if (pval.ItemUID == "Btn02")
							{
								LoadData();
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							if (pval.CharPressed == 9)
							{
								if (pval.ItemUID == "CntcCode")
								{
									if (string.IsNullOrEmpty(oForm.Items.Item(pval.ItemUID).Specific.Value))
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
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:							//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:						//17
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:							//1
							if (pval.ItemUID == "Btn01")
							{
								Print_Report01();
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							    //2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:							//5
							if (pval.ItemChanged == true)
							{
								oForm.Freeze(true);
								if (pval.ItemUID == "BPLId" || pval.ItemUID == "DocType")
								{
									oMat01.Clear();
									oDS_PS_FI420L.Clear();
								}
								oForm.Freeze(false);
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							        //6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:							//7
							if (pval.ItemUID == "Mat01" && pval.Row == Convert.ToDouble("0") && pval.ColUID == "Check")
							{
								oForm.Freeze(true);
								oMat01.FlushToDataSource();
								if ( string.IsNullOrEmpty(oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim()) || oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim() == "N")
								{
									Check = "Y";
								}
								else if (oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim() == "Y")
								{
									Check = "N";
								}
								for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
								{
									oDS_PS_FI420L.SetValue("U_ColReg01", i, Check);
								}
								oMat01.LoadFromDataSource();
								oForm.Freeze(false);
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:					//8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:							    //10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:							//11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:							//18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:						//19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:							//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:						//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:							    //3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:							    //4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                            //17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01); //메모리 해제
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI420L); //메모리 해제
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
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			SAPbouiCOM.OptionBtn optBtn = null;

			try
			{
				oDS_PS_FI420L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				oForm.DataSources.UserDataSources.Add("PntGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("PntGbn").Specific.DataBind.SetBound(true, "", "PntGbn");

				oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DocType").Specific.DataBind.SetBound(true, "", "DocType");

				oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

				oForm.DataSources.UserDataSources.Add("OptionDS01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				optBtn = oForm.Items.Item("Rad01").Specific;
				optBtn.ValOn = "1";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");

				//optBtn.Selected = True

				optBtn = oForm.Items.Item("Rad02").Specific;
				optBtn.ValOn = "2";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad03").Specific;
				optBtn.ValOn = "3";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad04").Specific;
				optBtn.ValOn = "4";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad05").Specific;
				optBtn.ValOn = "5";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				oForm.DataSources.UserDataSources.Add("OptionDS11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				optBtn = oForm.Items.Item("Rad11").Specific;
				optBtn.ValOn = "1";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");

				//optBtn.Selected = True

				optBtn = oForm.Items.Item("Rad12").Specific;
				optBtn.ValOn = "2";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad13").Specific;
				optBtn.ValOn = "3";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad14").Specific;
				optBtn.ValOn = "4";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad15").Specific;
				optBtn.ValOn = "5";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(optBtn);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		public void ComboBox_Setting()
		{
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//콤보에 기본값설정
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				// 전표유형
				oForm.Items.Item("DocType").Specific.ValidValues.Add("24", "입금");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("46", "지급");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("13", "판매");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("99", "기타(입금,지급,판매,제외)");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("00", "전체");
				oForm.Items.Item("DocType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("PntGbn").Specific.ValidValues.Add("10", "연결발행");
				oForm.Items.Item("PntGbn").Specific.ValidValues.Add("20", "개별발행");
				oForm.Items.Item("PntGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
		public void Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
			}
		}

		/// <summary>
		/// LoadData
		/// </summary>
		public void LoadData()
		{
			int i = 0;
			string sQry = String.Empty;

			
			string BPLID = String.Empty;
			//System.DateTime DocDate = default(System.DateTime);
			string DocDate = String.Empty;
			string DocType = String.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);
			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();
				//DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

				if (String.IsNullOrEmpty(DocDate))
				{
					PSH_Globals.SBO_Application.MessageBox("전기일자는 필수입력사항 입니다. 확인하세요.");
					return;
				}

				sQry = "EXEC [PS_FI420_01] '" + BPLID + "','" + DocType + "','" + DocDate + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_FI420L.Clear();

                if (oRecordSet.RecordCount == 0)
                {
                    oForm.Freeze(true);
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oForm.Freeze(false);
                    return;
                }

                oForm.Freeze(true);

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_FI420L.Size)
					{
						oDS_PS_FI420L.InsertRecord((i));
					}

					oMat01.AddRow();
					oDS_PS_FI420L.Offset = i;
					oDS_PS_FI420L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_FI420L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd"));  //  날짜형식으로 Convert
					oDS_PS_FI420L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDueDate").Value.ToString().Trim()).ToString("yyyyMMdd"));  //  날짜형식으로 Convert
					oDS_PS_FI420L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("DocTotal").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("JrnlMemo").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("TransId").Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgBar01.Value = ProgBar01.Value + 1;
					ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgBar01.Stop();
				oForm.Freeze(false);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
			}
		}

		/// <summary>
		/// Print_Report01
		/// </summary>
		private void Print_Report01()
		{
			int i = 0;
			int ErrNum = 0;
			string WinTitle = String.Empty;
			string ReportName = String.Empty;
			string DocType = String.Empty;
			string sQry = String.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();

				WinTitle = "회계전표 [PS_FI420]";

				if (oForm.Items.Item("PntGbn").Specific.Value.ToString().Trim() == "20")
				{
					ReportName = "PS_FI420_02.rpt";
				}
				else
				{
					ReportName = "PS_FI420_01.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@RadBtn01", oForm.DataSources.UserDataSources.Item("OptionDS01").Value));
				dataPackFormula.Add(new PSH_DataPackClass("@RadBtn11", oForm.DataSources.UserDataSources.Item("OptionDS11").Value));

				sQry = "Delete [Z_PS_FI420]";
				oRecordSet.DoQuery(sQry);

				oMat01.FlushToDataSource();
				for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
				{
					if (oDS_PS_FI420L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						sQry = "Insert [Z_PS_FI420] values ('" + oDS_PS_FI420L.GetValue("U_ColReg06", i).ToString().Trim() + "')";
						oRecordSet.DoQuery(sQry);
					}
				}

				// 조회조건문

				//// 조회조건문  (원본)
				//sQry = "EXEC [PS_FI420_02] '" + oForm.Items.Item("DocType").Specific.Value.ToString().Trim() + "'";
				//oRecordSet.DoQuery(sQry);
				////    If oRecordSet01.RecordCount = 0 Then
				////        ErrNum = 1
				////        GoTo Print_Report01_Error
				////    End If
				//if (oForm.Items.Item("DocType").Specific.Value.ToString().Trim() == "13")
				//{
				//	sQry = " Select * From  ZPS_FI420_TEMP Order by U_RptItm01,TransId, Convert(Numeric(12,0),Line_Id)";
				//}
				//else
				//{
				//	sQry = "Select  * From  ZPS_FI420_TEMP Order by TransId, Convert(Numeric(12,0),Line_Id) ";
				//}

				// 마이그레션시 FI420_02로 통합해서 새로작성  2020.09.21

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocType", DocType));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}

			

			//oRecordSet.DoQuery(sQry);
			//if (oRecordSet.RecordCount == 0)
			//{
			//	ErrNum = 1;
			//	goto Print_Report01_Error;
			//}
			////
			//////CR Action
			//if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") == false)
			//{
			//	SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
			//}

			////UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			//oRecordSet01 = null;
		}
	}
}
