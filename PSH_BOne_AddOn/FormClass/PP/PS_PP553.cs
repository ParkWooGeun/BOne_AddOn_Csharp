using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 공정진행현황 조회
	/// </summary>
	internal class PS_PP553 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DataTable oDS_PS_PP553A;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP553.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP553_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP553");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP553_CreateItems();
				PS_PP553_ComboBox_Setting();
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
				oForm.ActiveItem = "BPLId"; //최초 커서위치
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP553_CreateItems
		/// </summary>
		private void PS_PP553_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_PP553A");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP553A");
				oDS_PS_PP553A = oForm.DataSources.DataTables.Item("PS_PP553A");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기간(공정대기)(시작)
				oForm.DataSources.UserDataSources.Add("FrDt1", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt1").Specific.DataBind.SetBound(true, "", "FrDt1");

				//기간(공정대기)(종료)
				oForm.DataSources.UserDataSources.Add("ToDt1", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt1").Specific.DataBind.SetBound(true, "", "ToDt1");

				//기간(공정시작)(시작)
				oForm.DataSources.UserDataSources.Add("FrDt2", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt2").Specific.DataBind.SetBound(true, "", "FrDt2");

				//기간(공정시작)(종료)
				oForm.DataSources.UserDataSources.Add("ToDt2", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt2").Specific.DataBind.SetBound(true, "", "ToDt2");

				//기간(공정완료)(시작)
				oForm.DataSources.UserDataSources.Add("FrDt3", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt3").Specific.DataBind.SetBound(true, "", "FrDt3");

				//기간(공정완료)(종료)
				oForm.DataSources.UserDataSources.Add("ToDt3", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt3").Specific.DataBind.SetBound(true, "", "ToDt3");

				//작업구분
				oForm.DataSources.UserDataSources.Add("WorkGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("WorkGbn").Specific.DataBind.SetBound(true, "", "WorkGbn");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("SubNo1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("SubNo1").Specific.DataBind.SetBound(true, "", "SubNo1");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("SubNo2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("SubNo2").Specific.DataBind.SetBound(true, "", "SubNo2");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//공정코드
				oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

				//공정명
				oForm.DataSources.UserDataSources.Add("CpName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CpName").Specific.DataBind.SetBound(true, "", "CpName");

				//등록자사번
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//등록자성명
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP553_ComboBox_Setting
		/// </summary>
		private void PS_PP553_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", dataHelpClass.User_BPLID(), false, false);

				//작업구분_S
				sQry = " SELECT      Code, ";
				sQry += "                Name ";
				sQry += " FROM       [@PSH_ITMBSORT]";
				sQry += " WHERE      U_PudYN = 'Y'";
				sQry += " ORDER BY  Code";
				oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn").Specific, sQry, "%", false, false);

				//거래처구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP553_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP553_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;

			string OrdNum;
			string SubNo1;
			string SubNo2;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CpCode": //공정
						oForm.Items.Item("CpName").Specific.VALUE = dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'", "");
						break;

					case "CntcCode": //등록자
						oForm.Items.Item("CntcName").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
						break;
				}

				if (oUID == "OrdNum" || oUID == "SubNo1" || oUID == "SubNo2")
				{
					OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
					SubNo1 = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
					SubNo2 = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();

					sQry = "  SELECT   CASE";
					sQry += "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "                 ELSE T0.U_JakMyung";
					sQry += "             END AS [ItemName],";
					sQry += "             CASE";
					sQry += "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
					sQry += "                 ELSE T0.U_JakSize";
					sQry += "             END AS [ItemSpec]";
					sQry += " FROM     [@PS_PP020H] AS T0";
					sQry += " WHERE   T0.U_JakName = '" + OrdNum + "'";
					sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + SubNo1 + "' = '' THEN '00' ELSE '" + SubNo1 + "' END";
					sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + SubNo2 + "' = '' THEN '000' ELSE '" + SubNo2 + "' END";

					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("ItemName").Specific.VALUE = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
					oForm.Items.Item("ItemSpec").Specific.VALUE = oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim();
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
		/// PS_PP553_MTX01
		/// </summary>
		private void PS_PP553_MTX01()
		{
			string sQry;
			string errMessage = String.Empty;

			string BPLID;	 //사업장
			string FrDt1;	 //기간(Fr)
			string ToDt1;	 //기간(To)
			string FrDt2;	 //기간(Fr)
			string ToDt2;	 //기간(To)
			string FrDt3;	 //기간(Fr)
			string ToDt3;	 //기간(To)
			string WorkGbn;	 //작업구분
			string OrdNum;	 //작번
			string SubNo1;	 //서브작번1
			string SubNo2;	 //서브작번2
			string CpCode;	 //공정
			string CntcCode; //등록자사번
			string CardType; //거래처구분
			string ItemType; //품목구분

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID    = oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim();
				FrDt1    = oForm.Items.Item("FrDt1").Specific.Value.ToString().Trim();
				ToDt1    = oForm.Items.Item("ToDt1").Specific.Value.ToString().Trim();
				FrDt2    = oForm.Items.Item("FrDt2").Specific.Value.ToString().Trim();
				ToDt2    = oForm.Items.Item("ToDt2").Specific.Value.ToString().Trim();
				FrDt3    = oForm.Items.Item("FrDt3").Specific.Value.ToString().Trim();
				ToDt3    = oForm.Items.Item("ToDt3").Specific.Value.ToString().Trim();
				WorkGbn  = oForm.Items.Item("WorkGbn").Specific.Selected.Value.ToString().Trim();
				OrdNum   = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SubNo1   = oForm.Items.Item("SubNo1").Specific.Value.ToString().Trim();
				SubNo2   = oForm.Items.Item("SubNo2").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
				CpCode   = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP553_01 '";
				sQry += BPLID + "','";
				sQry += FrDt1 + "','";
				sQry += ToDt1 + "','";
				sQry += FrDt2 + "','";
				sQry += ToDt2 + "','";
				sQry += FrDt3 + "','";
				sQry += ToDt3 + "','";
				sQry += WorkGbn + "','";
				sQry += OrdNum + "','";
				sQry += SubNo1 + "','";
				sQry += SubNo2 + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += CpCode + "','";
				sQry += CntcCode + "'";

				oGrid.DataTable.Clear();
				oDS_PS_PP553A.ExecuteQuery(sQry);

				oGrid.Columns.Item(7).RightJustified = true;
				oGrid.Columns.Item(14).RightJustified = true;

				oGrid.Columns.Item("대기-시작소요일").RightJustified = true;
				oGrid.Columns.Item("시작-완료소요일").RightJustified = true;

				oGrid.Columns.Item("대기등록시간").Visible = false;
				oGrid.Columns.Item("시작등록시간").Visible = false;
				oGrid.Columns.Item("완료등록시간").Visible = false;

				if (oGrid.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oGrid.AutoResizeColumns();
				oForm.Update();
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
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
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP553_MTX01();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP553_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
				}
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
						PS_PP553_FlushToItemValue(pVal.ItemUID, 0, "");
					}
				}
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (oGrid.Columns.Count > 0)
					{
						oGrid.AutoResizeColumns();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						case "1285": //복원
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
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
						case "1285": //복원
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
