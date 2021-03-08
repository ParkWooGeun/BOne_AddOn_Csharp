using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 파카 라벨 출력
	/// </summary>
	internal class PS_SD012 : PSH_BaseClass
	{

		public string oFormUniqueID01;
		private String oLast_Item_UID;

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD012.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_SD012_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_SD012");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				CreateItems();

				oForm.EnableMenu(("1281"), false);			// 찾기
				oForm.EnableMenu(("1282"), true);			// 추가
				oForm.EnableMenu(("1283"), true);			// 제거
				oForm.EnableMenu(("1287"), true);			// 복제
				oForm.EnableMenu(("1284"), false);			// 취소
				oForm.EnableMenu(("1288"), false);
				oForm.EnableMenu(("1289"), false);
				oForm.EnableMenu(("1290"), false);
				oForm.EnableMenu(("1291"), false);
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
            try {
				if ((pval.BeforeAction == true)) {
					switch (pval.EventType) {
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                   //1
							if (pval.ItemUID == "BtnPrint")
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
								{
									if (HeaderSpaceLineDel() == false)
									{
										BubbleEvent = false;
										return;
									} else
									{
										PS_SD012_Print_Report();
									}
								}
							}
							break;

						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                       //2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                   //5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:                          //6
							break;
						case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                   //7
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:            //8
							break;
						case SAPbouiCOM.BoEventTypes.et_VALIDATE:                       //10
							break;
						case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                    //11
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                  //18
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                //19
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                    //20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:               //27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                      //3
							oLast_Item_UID = pval.ItemUID;
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                     //4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
							break;
					}
				} else if ((pval.BeforeAction == false)) {
					switch (pval.EventType) {
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
							break;
						case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
							break;
						case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
							break;
						case SAPbouiCOM.BoEventTypes.et_CLICK:							//6
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
						case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					//20
							break;
						case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
							break;
						case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:						//3
							oLast_Item_UID = pval.ItemUID;
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
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
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
						case "1293":							//행삭제
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
						case "1281":							//찾기
							break;
						case "1282":							//추가
							FormItem_Clear();
							break;
						case "1287":							//복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
						case "1293":							//행삭제
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("PartNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8);
				oForm.Items.Item("PartNo").Specific.DataBind.SetBound(true, "", "PartNo");

				oForm.DataSources.UserDataSources.Add("ModelNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("ModelNo").Specific.DataBind.SetBound(true, "", "ModelNo");

				oForm.DataSources.UserDataSources.Add("DataCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("DataCode").Specific.DataBind.SetBound(true, "", "DataCode");
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
		/// FormItemEnabled
		/// </summary>
		public void FormItemEnabled()
		{
			try
			{
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{

				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
				{
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
			}
		}

		/// <summary>
		/// FormItem_Clear
		/// </summary>
		public void FormItem_Clear()
		{
			try
			{
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
				{
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
			}
		}

		/// <summary>
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			short ErrNum = 0;
			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("PartNo").Specific.VALUE.ToString().Trim()))
				{
					ErrNum = 1;
					throw new Exception();
				}
						
				if (string.IsNullOrEmpty(oForm.Items.Item("ModelNo").Specific.VALUE.ToString().Trim()))
				{
					ErrNum = 2;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DataCode").Specific.VALUE.ToString().Trim()))
				{
					ErrNum = 3;
					throw new Exception();
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("PartNo를 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("ModelNo를 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("DataCode를 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// PS_SD012_Print_Report
		/// </summary>
		private void PS_SD012_Print_Report()
		{
			string WinTitle;
			string ReportName;

			string PartNo;			//PartNo
			string ModelNo;			//ModelNo
			string DataCode;         //DataCode

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				// 조회조건문
				PartNo = oForm.Items.Item("PartNo").Specific.Value.ToString().Trim();
				ModelNo = oForm.Items.Item("ModelNo").Specific.Value.ToString().Trim();
				DataCode = oForm.Items.Item("DataCode").Specific.Value.ToString().Trim();

				// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
				WinTitle = "[PS_SD012] 레포트";
				ReportName = "PS_SD012_01.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@PartNo", PartNo));
				dataPackParameter.Add(new PSH_DataPackClass("@ModelNo", ModelNo));
				dataPackParameter.Add(new PSH_DataPackClass("@DataCode", DataCode));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);

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
