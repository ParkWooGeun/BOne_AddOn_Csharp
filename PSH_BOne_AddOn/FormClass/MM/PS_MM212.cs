using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 구매진도현황
	/// </summary>
	internal class PS_MM212 : PSH_BaseClass
	{
		private string oFormUniqueID;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM212.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM212_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM212");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;

				oForm.Freeze(true);
				PS_MM212_CreateItems();
				PS_MM212_ComboBox_Setting();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM212_CreateItems
		/// </summary>
		private void PS_MM212_CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");

				//품의여부
				oForm.DataSources.UserDataSources.Add("MM030YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
				oForm.Items.Item("MM030YN").Specific.DataBind.SetBound(true, "", "MM030YN");

				//입고여부
				oForm.DataSources.UserDataSources.Add("MM050YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
				oForm.Items.Item("MM050YN").Specific.DataBind.SetBound(true, "", "MM050YN");

				//검수여부
				oForm.DataSources.UserDataSources.Add("MM070YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
				oForm.Items.Item("MM070YN").Specific.DataBind.SetBound(true, "", "MM070YN");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM212_ComboBox_Setting
		/// </summary>
		private void PS_MM212_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "SELECT U_Minor, U_CdName  From [@PS_SY001L] WHERE Code = 'C105' AND U_UseYN Like '%' ORDER BY U_Seq";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//구분
				oForm.Items.Item("Gbn01").Specific.ValidValues.Add("10", "청구기준(청구자별)");
				oForm.Items.Item("Gbn01").Specific.ValidValues.Add("20", "품의기준(거래처별)");
				oForm.Items.Item("Gbn01").Specific.ValidValues.Add("30", "청구기준(작번별)");
				oForm.Items.Item("Gbn01").Specific.ValidValues.Add("40", "청구기준(분류별)");
				oForm.Items.Item("Gbn01").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//구매방식
				sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("Purchase").Specific.ValidValues.Add("", "");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("Purchase").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//품목대분류
				sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("ItmBSort").Specific.ValidValues.Add("", "");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//사용처
				sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("UseDept").Specific.ValidValues.Add("", "");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("UseDept").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//품의여부
				oForm.Items.Item("MM030YN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("MM030YN").Specific.ValidValues.Add("Y", "완료");
				oForm.Items.Item("MM030YN").Specific.ValidValues.Add("N", "미완료");
				oForm.Items.Item("MM030YN").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//입고여부
				oForm.Items.Item("MM050YN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("MM050YN").Specific.ValidValues.Add("Y", "완료");
				oForm.Items.Item("MM050YN").Specific.ValidValues.Add("N", "미완료");
				oForm.Items.Item("MM050YN").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//검수여부
				oForm.Items.Item("MM070YN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("MM070YN").Specific.ValidValues.Add("Y", "완료");
				oForm.Items.Item("MM070YN").Specific.ValidValues.Add("N", "미완료");
				oForm.Items.Item("MM070YN").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//아이디별 사번 세팅
				oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM212_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_MM212_FlushToItemValue(string oUID)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "ItemCode":
						sQry = "Select FrgnName, U_Size From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						oForm.Items.Item("Size").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
						break;
					case "CntcCode":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
					case "CardCode":
						sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM212_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM212_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim()))
				{
					errMessage = "기준년도는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("Gbn01").Specific.Value.ToString().Trim() == "30" && string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim()))
				{
					errMessage = "작번별 구매진도현황은 메인작번을 필수로 입력해야 합니다. 확인하세요.";
					throw new Exception();
				}

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_MM212_Print_Report01 리포트 출력
		/// </summary>
		[STAThread]
		private void PS_MM212_Print_Report01()
		{
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string BPLId;
			string Gbn01;
			string Purchase;
			string ItmBsort;
			string ItemCode;
			string DocDateFr;
			string DocDateTo;
			string CntcCode;
			string CardCode;
			string OrdNum;
			string OrdSub1;
			string OrdSub2;
			string UseDept;
			string ItemName;
			string Size;
			string MM030YN;
			string MM050YN;
			string MM070YN;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Gbn01 = oForm.Items.Item("Gbn01").Specific.Value.ToString().Trim();
				Purchase = oForm.Items.Item("Purchase").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				OrdSub1 = oForm.Items.Item("OrdSub1").Specific.Value.ToString().Trim();
				OrdSub2 = oForm.Items.Item("OrdSub2").Specific.Value.ToString().Trim();
				UseDept = oForm.Items.Item("UseDept").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim().Replace("'", "''");
				Size = oForm.Items.Item("Size").Specific.Value.ToString().Trim().Replace("'", "''");
				MM030YN = oForm.Items.Item("MM030YN").Specific.Value.ToString().Trim();
				MM050YN = oForm.Items.Item("MM050YN").Specific.Value.ToString().Trim();
				MM070YN = oForm.Items.Item("MM070YN").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(BPLId))
				{
					BPLId = "%";
				}
				if (string.IsNullOrEmpty(Purchase))
				{
					Purchase = "%";
				}
				if (string.IsNullOrEmpty(ItmBsort))
				{
					ItmBsort = "%";
				}
				if (string.IsNullOrEmpty(ItemCode))
				{
					ItemCode = "%";
				}
				if (string.IsNullOrEmpty(DocDateFr))
				{
					DocDateFr = "18990101";
				}
				if (string.IsNullOrEmpty(DocDateTo))
				{
					DocDateTo = "20991231";
				}
				if (string.IsNullOrEmpty(CntcCode))
				{
					CntcCode = "%";
				}
				if (string.IsNullOrEmpty(CardCode))
				{
					CardCode = "%";
				}
				if (string.IsNullOrEmpty(OrdNum))
				{
					OrdNum = "%";
				}
				if (string.IsNullOrEmpty(OrdSub1))
				{
					OrdSub1 = "%";
				}
				if (string.IsNullOrEmpty(OrdSub2))
				{
					OrdSub2 = "%";
				}
				if (string.IsNullOrEmpty(UseDept))
				{
					UseDept = "%";
				}

				if (Gbn01 == "10")
				{
					WinTitle = "구매 진도 현황 [PS_MM212_01]";
					ReportName = "PS_MM212_01.rpt";
				}
				else if (Gbn01 == "20")
				{
					WinTitle = "구매 진도 현황 [PS_MM212_02]";
					ReportName = "PS_MM212_02.rpt";
				}
				else if (Gbn01 == "30")
				{
					WinTitle = "구매 진도 현황 [PS_MM212_03]";
					ReportName = "PS_MM212_03.rpt";
				}
				else if (Gbn01 == "40")
				{
					WinTitle = "구매 진도 현황 [PS_MM212_04]";
					ReportName = "PS_MM212_04.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				//Formula List
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLId, "")));
				dataPackFormula.Add(new PSH_DataPackClass("@F01", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@F02", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2)));

				if (Purchase == "%")
				{
					dataPackFormula.Add(new PSH_DataPackClass("@F03", "전체"));
				}
				else
				{
					dataPackFormula.Add(new PSH_DataPackClass("@F03", dataHelpClass.Get_ReData("Name", "Code", "[@PSH_ORDTYP]", Purchase, "")));
				}

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@Gbn1", Gbn01));
				dataPackParameter.Add(new PSH_DataPackClass("@Purchase", Purchase));
				dataPackParameter.Add(new PSH_DataPackClass("@ItmBsort", ItmBsort));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr));
				dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo));
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdSub1", OrdSub1));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdSub2", OrdSub2));
				dataPackParameter.Add(new PSH_DataPackClass("@UseDept", UseDept));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemName", ItemName));
				dataPackParameter.Add(new PSH_DataPackClass("@Size", Size));
				dataPackParameter.Add(new PSH_DataPackClass("@MM030YN", MM030YN));
				dataPackParameter.Add(new PSH_DataPackClass("@MM050YN", MM050YN));
				dataPackParameter.Add(new PSH_DataPackClass("@MM070YN", MM070YN));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			switch (pVal.EventType)
			{
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
				//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
				//    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
				//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
				//    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
				//    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
				//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_Drag: //39
				//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
				//    break;
			}
		}

		/// <summary>
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (PS_MM212_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}

						System.Threading.Thread thread = new System.Threading.Thread(PS_MM212_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "CntcCode" || pVal.ItemUID == "CardCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// COMBO_SELECT 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Gbn01")
					{
						if (oForm.Items.Item("Gbn01").Specific.Value == "10")
						{
							oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
							oForm.Items.Item("MM030YN").Enabled = true;
						}
						else if (oForm.Items.Item("Gbn01").Specific.Value == "20")
						{
							oForm.Items.Item("CntcCode").Specific.Value = "";
							oForm.Items.Item("CntcName").Specific.Value = "";
							oForm.Items.Item("MM030YN").Enabled = false;
							oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (oForm.Items.Item("Gbn01").Specific.Value == "30")
						{
							oForm.Items.Item("CntcCode").Specific.Value = "";
							oForm.Items.Item("CntcName").Specific.Value = "";
							oForm.Items.Item("MM030YN").Enabled = true;
							oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							oForm.Items.Item("CntcCode").Specific.Value = "";
							oForm.Items.Item("CntcName").Specific.Value = "";
							oForm.Items.Item("MM030YN").Enabled = true;
							oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "CntcCode" || pVal.ItemUID == "CardCode")
						{
							PS_MM212_FlushToItemValue(pVal.ItemUID);
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{ 
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// FORM_UNLOAD 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// FormDataEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="BusinessObjectInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
		{
			try
			{
				switch (BusinessObjectInfo.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
