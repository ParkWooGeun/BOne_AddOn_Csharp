using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 설비종합효율-기초정보조회
	/// </summary>
	internal class PS_PP915 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;

		private SAPbouiCOM.DataTable oDS_PS_PP915A;
		private SAPbouiCOM.DataTable oDS_PS_PP915B;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP915.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP915_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP915");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP915_CreateItems();
				PS_PP915_ComboBox_Setting();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP915_CreateItems
		/// </summary>
		private void PS_PP915_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;

				oForm.DataSources.DataTables.Add("PS_PP915A");
				oForm.DataSources.DataTables.Add("PS_PP915B");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP915A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP915B");

				oDS_PS_PP915A = oForm.DataSources.DataTables.Item("PS_PP915A");
				oDS_PS_PP915B = oForm.DataSources.DataTables.Item("PS_PP915B");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//작업구분
				oForm.DataSources.UserDataSources.Add("OrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OrdGbn").Specific.DataBind.SetBound(true, "", "OrdGbn");

				//기준년월(시작)
				oForm.DataSources.UserDataSources.Add("FrYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("FrYM").Specific.DataBind.SetBound(true, "", "FrYM");
				oForm.Items.Item("FrYM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

				//기준년월(종료)
				oForm.DataSources.UserDataSources.Add("ToYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("ToYM").Specific.DataBind.SetBound(true, "", "ToYM");
				oForm.Items.Item("ToYM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

				//장비코드
				oForm.DataSources.UserDataSources.Add("MachCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("MachCode").Specific.DataBind.SetBound(true, "", "MachCode");

				//장비명
				oForm.DataSources.UserDataSources.Add("MachName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("MachName").Specific.DataBind.SetBound(true, "", "MachName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP915_ComboBox_Setting
		/// </summary>
		private void PS_PP915_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", "1", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//작업구분_S
				sQry = " SELECT      U_Minor,";
				sQry += "                 U_CdName";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'P205'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP915_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP915_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "MachCode": //장비명
						oForm.Items.Item("MachName").Specific.Value = dataHelpClass.Get_ReData("U_MachName", "U_MachCode", "[@PS_PP130H]", "'" + oForm.Items.Item("MachCode").Specific.Value.ToString().Trim() + "'", "");
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP915_FormResize
		/// </summary>
		private void PS_PP915_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 30;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP915_MTX01
		/// </summary>
		private void PS_PP915_MTX01()
		{
			string sQry;
			string errMessage = String.Empty;

			string BPLID;   //사업장
			string OrdGbn;  //작업구분
			string FrYM;    //기준년월(시작)
			string ToYM;    //기준년월(종료)
			string MachCode;//장비코드

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim();
				OrdGbn = oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim();
				FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
				ToYM = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
				MachCode = oForm.Items.Item("MachCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				//작업일보 조회
				sQry = "    EXEC PS_PP915_01 '";
				sQry += BPLID + "','";
				sQry += OrdGbn + "','";
				sQry += FrYM + "','";
				sQry += ToYM + "','";
				sQry += MachCode + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP915A.ExecuteQuery(sQry);

				oGrid01.Columns.Item(12).RightJustified = true;
				oGrid01.Columns.Item(13).RightJustified = true;
				oGrid01.Columns.Item(14).RightJustified = true;
				oGrid01.Columns.Item(17).RightJustified = true;
				oGrid01.Columns.Item(20).RightJustified = true;
				oGrid01.Columns.Item(24).RightJustified = true;
				oGrid01.Columns.Item(27).RightJustified = true;
				oGrid01.Columns.Item(28).RightJustified = true;
				oGrid01.Columns.Item(29).RightJustified = true;
				oGrid01.Columns.Item(30).RightJustified = true;
				oGrid01.Columns.Item(31).RightJustified = true;
				oGrid01.Columns.Item(32).RightJustified = true;
				oGrid01.Columns.Item(33).RightJustified = true;
				oGrid01.Columns.Item(34).RightJustified = true;
				oGrid01.Columns.Item(35).RightJustified = true;
				oGrid01.Columns.Item(36).RightJustified = true;
				oGrid01.Columns.Item(37).RightJustified = true;
				oGrid01.Columns.Item(38).RightJustified = true;
				oGrid01.Columns.Item(39).RightJustified = true;
				oGrid01.Columns.Item(40).RightJustified = true;
				oGrid01.Columns.Item(41).RightJustified = true;
				oGrid01.Columns.Item(42).RightJustified = true;
				oGrid01.Columns.Item(43).RightJustified = true;
				oGrid01.Columns.Item(44).RightJustified = true;
				oGrid01.Columns.Item(45).RightJustified = true;
				oGrid01.Columns.Item(46).RightJustified = true;
				oGrid01.Columns.Item(47).RightJustified = true;
				oGrid01.Columns.Item(48).RightJustified = true;
				oGrid01.Columns.Item(49).RightJustified = true;
				oGrid01.Columns.Item(50).RightJustified = true;
				oGrid01.Columns.Item(51).RightJustified = true;
				oGrid01.Columns.Item(52).RightJustified = true;
				oGrid01.Columns.Item(53).RightJustified = true;
				oGrid01.Columns.Item(54).RightJustified = true;
				oGrid01.Columns.Item(55).RightJustified = true;

				//장비가동시간
				sQry = "    EXEC PS_PP915_02 '";
				sQry += BPLID + "','";
				sQry += OrdGbn + "','";
				sQry += FrYM + "','";
				sQry += ToYM + "','";
				sQry += MachCode + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP915B.ExecuteQuery(sQry);

				oGrid02.Columns.Item(8).RightJustified = true;
				oGrid02.Columns.Item(11).RightJustified = true;
				oGrid02.Columns.Item(14).RightJustified = true;
				oGrid02.Columns.Item(17).RightJustified = true;
				oGrid02.Columns.Item(20).RightJustified = true;

				if (oGrid01.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid01.AutoResizeColumns();
				oForm.Update();
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
							PS_PP915_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
					}
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MachCode", "");
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
					PS_PP915_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
				{
					PS_PP915_FormResize();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					PS_PP915_FlushToItemValue(pVal.ItemUID, 0, "");
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP915A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP915B);
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
