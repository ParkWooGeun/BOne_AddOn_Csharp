using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 영업외손익.특별손익명세서 PS_FI150
	/// </summary>
	internal class PS_FI150 : PSH_BaseClass
	{
		private string oFormUniqueID01;

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI150.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI150_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI150");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
		
				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();

				oForm.EnableMenu("1283", false);				// 삭제
				oForm.EnableMenu("1286", false);				// 닫기
				oForm.EnableMenu("1287", false);				// 복제
				oForm.EnableMenu("1284", false);				// 취소
				oForm.EnableMenu("1293", false);				// 행삭제
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
				oForm.ActiveItem = "BPLId";   //최초 커서위치

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
         		oForm.DataSources.UserDataSources.Add("YYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("YYYYMM").Specific.DataBind.SetBound(true, "", "YYYYMM");
				oForm.DataSources.UserDataSources.Item("YYYYMM").Value = DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") ;
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
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
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

				oForm.Items.Item("YYYYMM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			int ErrNum = 0;
			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
				if (oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim().Length != 7)
				{
					ErrNum = 2;
					throw new Exception();
				}

				functionReturnValue = true;

			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("해당년월은 필수사항입니다. 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("해당년월(YYYY-MM)의 형식을 확인하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// Print_Query
		/// </summary>
		[STAThread]
		private void Print_Query()
		{
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string YYYYMM = string.Empty;
			string BPLId = string.Empty;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();

				WinTitle = "[PS_FI150] 영업외손익.특별손익명세서";
				ReportName = "PS_FI150_01.RPT";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula
				dataPackFormula.Add(new PSH_DataPackClass("@YYYYMM", YYYYMM));
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLId));

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@YYYYMM", YYYYMM));
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
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:	// 1
							if (pVal.ItemUID == "1")
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
								{
								}
							}
							else if (pVal.ItemUID == "Btn01")  //출력버튼 클릭시
							{
								if (HeaderSpaceLineDel() == false)
								{
									BubbleEvent = false;
									return;
								}
								else
								{
									System.Threading.Thread thread = new System.Threading.Thread(Print_Query);
									thread.SetApartmentState(System.Threading.ApartmentState.STA);
									thread.Start();
								}
							}
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						// 2
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                      // 3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                     // 4
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					// 5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							// 6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					// 7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			// 8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:					    // 10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					// 11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    // 17
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					// 18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				// 19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					// 20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				// 27
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:						// 1
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						    // 2
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                          // 3
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                         // 4
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:						// 5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							    // 6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:						// 7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:				// 8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:							// 10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:						// 11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                        // 17
							System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
							SubMain.Remove_Forms(oFormUniqueID01);
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:						// 18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:					// 19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:						// 20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:					// 27
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
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
                if ((eventInfo.BeforeAction == true))
				{
				}
				else if ((eventInfo.BeforeAction == false))
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
		/// Raise_FormDataEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="BusinessObjectInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
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
	}
}
