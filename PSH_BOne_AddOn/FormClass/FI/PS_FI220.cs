using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 합계잔액시산표
	/// </summary>
	internal class PS_FI220 : PSH_BaseClass
	{
		public string oFormUniqueID01;

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI220.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				oFormUniqueID01 = "PS_FI220_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI220");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;

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
			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                           //1
							if (pval.ItemUID == "Btn01")
							{
								if (HeaderSpaceLineDel() == false)
								{
									BubbleEvent = false;
									return;
								}
								Print_Report01();
							}
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                            //17
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
				else if ((BusinessObjectInfo.BeforeAction == false))
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
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
				if ((eventInfo.BeforeAction == true))
				{
					// 작업
				}
				else if ((eventInfo.BeforeAction == false))
				{
					// 작업
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
			try
			{
				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyy") + "-01-01";

				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("Level5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Level5").Specific.DataBind.SetBound(true, "", "Level5");
				oForm.DataSources.UserDataSources.Item("Level5").Value = "Y";

				oForm.DataSources.UserDataSources.Add("Level6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Level6").Specific.DataBind.SetBound(true, "", "Level6");
				oForm.DataSources.UserDataSources.Item("Level6").Value = "Y";
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
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
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
		/// Initialization
		/// </summary>
		public void Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 아이디별 사업장 세팅
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
		/// CF_ChooseFromList
		/// </summary>
		public void CF_ChooseFromList()
		{
			try
			{
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
				// Check
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim()))
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
					PSH_Globals.SBO_Application.StatusBar.SetText("기간은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// Print_Report01
		/// </summary>
		private void Print_Report01()
		{
			string WinTitle = String.Empty;
			string ReportName = String.Empty;
			
			string BPLId = String.Empty;
			string DocDateFr = String.Empty;
			string DocDateTo = String.Empty;
			string Level5 = String.Empty;
			string Level6 = String.Empty;

			string BPLName = String.Empty;
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{

				BPLId     = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				Level5    = oForm.DataSources.UserDataSources.Item("Level5").Value.ToString().Trim();
				Level6    = oForm.DataSources.UserDataSources.Item("Level6").Value.ToString().Trim();

				if (string.IsNullOrEmpty(BPLId))
				{
					BPLId = "%";
				}
				if (string.IsNullOrEmpty(DocDateFr))
				{
					DocDateFr = "18990101";
				}
				if (string.IsNullOrEmpty(DocDateTo))
				{
					DocDateTo = "20991231";
				}

				WinTitle = "합계잔액시산표 [PS_FI220_01]";
				ReportName = "PS_FI220_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				if (BPLId == "%")
				{
					BPLName = "전체";
				}
				else
				{
					sQry = "SELECT BPLName From [OBPL] Where BPLId = '" + BPLId + "'";
					oRecordSet.DoQuery(sQry);
					BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@F01", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@F02", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@F03", BPLName));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DateTime.ParseExact(DocDateFr, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DateTime.ParseExact(DocDateTo, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@Level5", Level5));
				dataPackParameter.Add(new PSH_DataPackClass("@Level6", Level6));

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
		}
	}
}
