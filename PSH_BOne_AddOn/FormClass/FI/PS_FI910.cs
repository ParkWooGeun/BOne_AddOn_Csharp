using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 매출채권 연령분석(받을어음)   PS_FI910
	/// </summary>
	internal class PS_FI910 : PSH_BaseClass
	{
		public string oFormUniqueID01;

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI910.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				// 매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI910_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI910");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();

				oForm.EnableMenu(("1283"), false);				// 삭제
				oForm.EnableMenu(("1286"), false);				// 닫기
				oForm.EnableMenu(("1287"), false);				// 복제
				oForm.EnableMenu(("1284"), false);				// 취소
				oForm.EnableMenu(("1293"), false);				// 행삭제
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
				oForm.DataSources.UserDataSources.Add("BaseDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("BaseDate").Specific.DataBind.SetBound(true, "", "BaseDate");
				oForm.DataSources.UserDataSources.Item("BaseDate").Value = DateTime.Now.ToString("yyyyMMdd");
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
		public void ComboBox_Setting()
		{
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 콤보에 기본값설정
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("0", "전체 사업장");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("BaseDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
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
								}
							}
							else if (pval.ItemUID == "Btn01") //출력버튼 클릭시
							{
								if (HeaderSpaceLineDel() == false)
								{
									BubbleEvent = false;
									return;
								}
								else
								{
									Print_Query();
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							if (pval.CharPressed == 9)
							{
								// 헤더
								if (pval.ItemUID == "SCardCode")
								{
									if (string.IsNullOrEmpty(oForm.Items.Item("SCardCode").Specific.VALUE))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
										BubbleEvent = false;
									}
								}
								if (pval.ItemUID == "ECardCode")
								{
									if (string.IsNullOrEmpty(oForm.Items.Item("ECardCode").Specific.VALUE))
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
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						//1
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:							//2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:						//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:						 	    //6
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
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        //17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
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
						case "1285":							//복원
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
						case "1285":							//복원
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
				}
				else if ((eventInfo.BeforeAction == false))
				{
					////작업
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
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
		{
			try
			{
				switch (oUID)
				{
				}
				if (oUID == "Mat01")
				{
					switch (oCol)
					{
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;

			int ErrNum = 0;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BaseDate").Specific.VALUE.ToString().Trim()))
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
					PSH_Globals.SBO_Application.StatusBar.SetText("기준일자는 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				functionReturnValue = false;
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		return functionReturnValue;
		}

		/// <summary>
		/// Print_Query
		/// </summary>
		private void Print_Query()
		{
			string WinTitle = String.Empty;
			string ReportName = String.Empty;

			string BaseDate = String.Empty;
			string SCardCode = String.Empty;
			string ECardCode = String.Empty;
			string BPLId = String.Empty;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BaseDate  = oForm.Items.Item("BaseDate").Specific.VALUE.ToString().Trim();
				SCardCode = oForm.Items.Item("SCardCode").Specific.VALUE.ToString().Trim();
				ECardCode = oForm.Items.Item("ECardCode").Specific.VALUE.ToString().Trim();
				BPLId     = oForm.Items.Item("BPLId").Specific.Selected.VALUE.ToString().Trim();

				if (string.IsNullOrEmpty(SCardCode))
				{
					SCardCode = "1";
				}
				if (string.IsNullOrEmpty(ECardCode))
				{
					ECardCode = "ZZZZZZZZ";
				}

				WinTitle = "[PS_FI910] 매출채권 연령분석(받을어음)";
				ReportName = "PS_FI910_01.RPT";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@BaseDate", BaseDate.Substring(0, 4) + "-" + BaseDate.Substring(4, 2) + "-" + BaseDate.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLId));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BaseDate", DateTime.ParseExact(BaseDate, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@SCardCode", SCardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ECardCode", ECardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));

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
