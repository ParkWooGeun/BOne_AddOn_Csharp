using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 설계실동공수등록현황(기계)
	/// </summary>
	internal class PS_PP909 : PSH_BaseClass
	{
		public string oFormUniqueID;

		public SAPbouiCOM.Grid oGrid;
		public SAPbouiCOM.Matrix oMat01;
		public SAPbouiCOM.Matrix oMat02;

		private SAPbouiCOM.DBDataSource oDS_PS_PP909L;  //라인(작업자정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP909M;  //라인(세부정보)

		public SAPbouiCOM.DataTable oDS_PS_PP909N;

		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP909.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP909_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP909");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP909_CreateItems();
				PS_PP909_ComboBox_Setting();
				PS_PP909_Initial_Setting();
				PS_PP909_FormResize();
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
				oForm.Items.Item("Folder02").Specific.Select(); //폼이 로드 될 때 Folder02이 선택됨
			}
		}

		/// <summary>
		/// PS_PP909_CreateItems
		/// </summary>
		private void PS_PP909_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_PP909L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP909M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				oGrid = oForm.Items.Item("Grid01").Specific;

				oForm.DataSources.DataTables.Add("PS_PP909N");

				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP909N");

				oDS_PS_PP909N = oForm.DataSources.DataTables.Item("PS_PP909N");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//작업자코드
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//작업자성명
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

				//작업일자Fr
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//작업일자To
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//사업장2
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//작업자코드2
				oForm.DataSources.UserDataSources.Add("CntcCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcCode02").Specific.DataBind.SetBound(true, "", "CntcCode02");

				//작업자성명2
				oForm.DataSources.UserDataSources.Add("CntcName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcName02").Specific.DataBind.SetBound(true, "", "CntcName02");

				//작업일자Fr2
				oForm.DataSources.UserDataSources.Add("FrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt02").Specific.DataBind.SetBound(true, "", "FrDt02");

				//작업일자To2
				oForm.DataSources.UserDataSources.Add("ToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt02").Specific.DataBind.SetBound(true, "", "ToDt02");
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
		/// PS_PP909_ComboBox_Setting
		/// </summary>
		public void PS_PP909_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName From [OBPL] order by 1", dataHelpClass.User_BPLID(), false, false);

				//사업장02
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName From [OBPL] order by 1", dataHelpClass.User_BPLID(), false, false);
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
		/// PS_PP909_Initial_Setting
		/// </summary>
		public void PS_PP909_Initial_Setting()
		{
			try
			{
				//날짜 설정
				oForm.Items.Item("FrDt").Specific.VALUE = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("FrDt02").Specific.VALUE = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt02").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

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
		/// PS_PP909_FormResize
		/// </summary>
		private void PS_PP909_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 90;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				oForm.Items.Item("Mat01").Width = oForm.Width / 2 - 90;
				oForm.Items.Item("Mat01").Height = oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height - oForm.Items.Item("2").Top) - 20;

				oForm.Items.Item("Mat02").Left = oForm.Width / 2 - 40;
				oForm.Items.Item("Mat02").Width = oForm.Width / 2 + 5;
				oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height;

				oForm.Items.Item("Static06").Left = oForm.Width / 2 - 40;

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
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
					Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
					Raise_EVENT_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "BtnSrch01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							//상세 정보 초기화
							oMat02.Clear();  
							oDS_PS_PP909M.Clear();

							PS_PP909_MTX01();			//매트릭스에 데이터 로드
						}
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						PS_PP909_MTX03();				//데이터 로드
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP909_Print_Report02);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스_S
					//Folder01이 선택되었을 때
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					//Folder02가 선택되었을 때
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "PS_PP909")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
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
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");                  //사원코드 포맷서치 활성
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode02", "");				//사원코드 포맷서치 활성
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;  //정렬
							oMat01.FlushToDataSource();
						}
						else
						{
							PS_PP909_MTX02(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE);
						}
					}
					else if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row == 0)
						{
							oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;   //정렬
							oMat02.FlushToDataSource();
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
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						PS_PP040 oTempClass = new PS_PP040();
						oTempClass.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.VALUE.ToString().Trim());
					}
					else if (pVal.ItemUID == "Mat02")
					{
						PS_PP030 oTempClass = new PS_PP030();
						oTempClass.LoadForm(oMat02.Columns.Item("PP030HNo").Cells.Item(pVal.Row).Specific.VALUE.ToString().Trim());
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CntcName").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();

						}
						else if (pVal.ItemUID == "CntcCode02")
						{
							sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CntcName02").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					PS_PP909_FormResize();
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
				else if (pVal.ItemUID == "Mat03")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat04")
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP909L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP909M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP909N);
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
						case "7169":							//엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							PS_PP909_AddMatrixRow(oMat01.VisualRowCount, false);
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
						case "7169":							//엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP909L.RemoveRecord(oDS_PS_PP909L.Size - 1);
							oMat01.LoadFromDataSource();
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
		/// PS_PP909_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		public void PS_PP909_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false)  //행추가여부
				{
					oDS_PS_PP909L.InsertRecord(oRow);
				}

				oMat01.AddRow();
				oDS_PS_PP909L.Offset = oRow;
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
		/// PS_PP909_MTX01
		/// </summary>
		private void PS_PP909_MTX01()
		{
			int ErrNum = 0;
			int loopCount;
			string sQry;

			string BPLID;			//사업장
			string CntcCode;		//작업자코드
			string FrDt;			//시작일자
			string ToDt;            //종료일자

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.VALUE.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt").Specific.VALUE.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt").Specific.VALUE.ToString().Trim();

				oForm.Freeze(true);

				sQry = "EXEC PS_PP909_01 '" + BPLID + "','" + FrDt + "','" + ToDt + "','" + CntcCode + "'";
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
						oDS_PS_PP909L.InsertRecord(loopCount);
					}
					oDS_PS_PP909L.Offset = loopCount;

					oDS_PS_PP909L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));					            //라인번호
					oDS_PS_PP909L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());	//작업일보문서번호
					oDS_PS_PP909L.SetValue("U_ColDt01", loopCount,  oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());		//작업일보등록일자
					oDS_PS_PP909L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim());	//사번
					oDS_PS_PP909L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());	//성명
					oDS_PS_PP909L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("YTime").Value.ToString().Trim());		//작업시간

					oRecordSet.MoveNext();
					ProgressBar01.Value = ProgressBar01.Value + 1;
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP909_MTX02
		/// </summary>
		/// <param name="prmDocEntry"></param>
		private void PS_PP909_MTX02(string prmDocEntry)
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

				sQry = "EXEC PS_PP909_02 '" + prmDocEntry + "'";
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
						oDS_PS_PP909M.InsertRecord(loopCount);
					}
					oDS_PS_PP909M.Offset = loopCount;

					oDS_PS_PP909M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));					                //라인번호
					oDS_PS_PP909M.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim());		//작업지시문서번호
					oDS_PS_PP909M.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim());		//작업지시라인번호
					oDS_PS_PP909M.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());			//메인작번
					oDS_PS_PP909M.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());			//서브작번1
					oDS_PS_PP909M.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());			//서브작번2
					oDS_PS_PP909M.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("WorkTime").Value.ToString().Trim());		//공정정보작업시간
					oDS_PS_PP909M.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("BdwQty").Value.ToString().Trim());			//기존도면매수
					oDS_PS_PP909M.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("DwRate").Value.ToString().Trim());			//적용율
					oDS_PS_PP909M.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("AdwQty").Value.ToString().Trim());			//적용매수
					oDS_PS_PP909M.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("NdwQty").Value.ToString().Trim());			//신규도면매수

					oRecordSet.MoveNext();
					ProgressBar01.Value = ProgressBar01.Value + 1;
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP909_MTX03
		/// </summary>
		private void PS_PP909_MTX03()
		{
			int ErrNum = 0;
			string sQry;

			string BPLID;           //사업장
			string CntcCode;        //작업자코드
			string FrDt;            //시작일자
			string ToDt;            //종료일자

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID02").Specific.VALUE.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.VALUE.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.VALUE.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.VALUE.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP909_03 '";
				sQry += BPLID + "','";
				sQry += CntcCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "'";

				oGrid.DataTable.Clear();
				oDS_PS_PP909N.ExecuteQuery(sQry);

				oGrid.Columns.Item(8).RightJustified = true;

				if (oGrid.Rows.Count == 1)
				{
					ErrNum = 1;
					throw new Exception();
				}

				oGrid.AutoResizeColumns();
				oForm.Update();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP909_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_PP909_Print_Report02()
		{
			string WinTitle;
			string ReportName;

			string BPLID;           //사업장
			string CntcCode;        //작업자코드
			string FrDt;            //시작일자
			string ToDt;            //종료일자

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID02").Specific.VALUE.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.VALUE.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt02").Specific.VALUE.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt02").Specific.VALUE.ToString().Trim();

				WinTitle = "[PS_PP909] 레포트";
				ReportName = "PS_PP909_02.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@WorkCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
