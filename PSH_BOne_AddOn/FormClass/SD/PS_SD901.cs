using System;
using SAPbouiCOM;
namespace PSH_BOne_AddOn
{
	/// <summary>
	/// (기계)판매오더 수주처 변경시 작업시시 변경처리
	/// </summary>
	internal class PS_SD901 : PSH_BaseClass
	{
		public string oFormUniqueID01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD901.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_SD901_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_SD901");
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString());
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
			try
			{
				if ((pval.BeforeAction == true))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                   //1
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
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                     //4
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
							break;
					}
				}
				else if ((pval.BeforeAction == false))
				{
					switch (pval.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                   //1
							if (pval.ItemUID == "Button01")
							{
								SAVE();
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
							break;
						case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                     //4
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
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
		/// ComboBox_Setting
		/// </summary>
		public void ComboBox_Setting()
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
		/// Initialization
		/// </summary>
		public void Initialization()
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

		private void SAVE()
		{
			int ErrNum = 0;
			string sQry;

			string ORDRNo;
			string ItemCode;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				ORDRNo = oForm.Items.Item("ORDRNo").Specific.VALUE.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.VALUE.ToString().Trim();

				if (string.IsNullOrEmpty(ORDRNo))
				{
					ErrNum = 1;
					throw new Exception();
				}

				if (string.IsNullOrEmpty(ItemCode))
				{
					ErrNum = 2;
					throw new Exception();
				}

				//판매오더, 작번 check
				sQry = "SELECT 'X' FROM [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry Inner join [OITM]  c On b.ItemCode = c.ItemCode And c.U_ItmBsort In ('105','106') WHERE a.DocEntry = '" + ORDRNo + "' And b.ItemCode = '" + ItemCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					ErrNum = 3;
					throw new Exception();
				}

				//조회조건문
				sQry = "EXEC [PS_SD901_01] '" + ORDRNo + "', '" + ItemCode + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.MessageBox("변경처리가 완료되었습니다.");
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("판매오더 번호가 없습니다.  확인해 주세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("품목코드(작지)번호가 없습니다.  확인해 주세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("판매오더가 없거나 판매오더에 해당작번이 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}
	}
}
