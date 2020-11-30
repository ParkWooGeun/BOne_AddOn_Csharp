using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월별 계정별 비용현황(상세)
	/// </summary>
	internal class PS_FI962 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Grid oGrid01;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="prmBPLId"></param>
		/// <param name="prmFrDt"></param>
		/// <param name="prmToDt"></param>
		/// <param name="prmAccount"></param>
		public void LoadForm(string prmBPLId, string prmFrDt, string prmToDt, string prmAccount)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI962.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI962_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI962");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_FI962_CreateItems();
				PS_FI962_MTX01(prmBPLId, prmFrDt, prmToDt, prmAccount);
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
		/// PS_FI962_CreateItems
		/// </summary>
		/// <returns></returns>
		private void PS_FI962_CreateItems()
		{
			try
			{
				oForm.Freeze(true);
				oGrid01 = oForm.Items.Item("Grid01").Specific;
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
		/// PS_FI962_MTX01
		/// </summary>
		/// <param name="prmBPLId"></param>
		/// <param name="prmFrDt"></param>
		/// <param name="prmToDt"></param>
		/// <param name="prmAccount"></param>
		private void PS_FI962_MTX01(string prmBPLId, string prmFrDt, string prmToDt, string prmAccount)
		{
			int ErrNum = 0;
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				sQry = "EXEC PS_FI961_03 '" + prmBPLId + "','" + prmFrDt + "','" + prmToDt + "','" + prmAccount + "'";

				oGrid01.DataTable.Clear();

				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				oGrid01.Columns.Item(2).RightJustified = true;              //금액 필드 Right 활성화

				if (oGrid01.Rows.Count == 0)
				{
					ErrNum = 1;
					throw new Exception();
				}
				oGrid01.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oForm.Freeze(false);
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
				switch (pVal.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
						//Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
						//Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
						//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:						    //6
						//Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
						//Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
						//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
						//Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
						//Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					//20
						//Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
						//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:						//3
						//Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:					//17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); //메모리 해제
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01); //메모리 해제
					SubMain.Remove_Forms(oFormUniqueID01);
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
