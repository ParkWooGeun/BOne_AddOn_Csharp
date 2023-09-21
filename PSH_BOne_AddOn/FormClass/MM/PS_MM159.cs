using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작지별 생산진행현황(사도급)
	/// </summary>
	internal class PS_MM159 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.Grid oGrid04;
		private SAPbouiCOM.Grid oGrid05;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM159.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM159_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM159");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_MM159_CreateItems();
				PS_MM159_SetComboBox();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
			}
		}

		/// <summary>
		/// PS_MM159_CreateItems
		/// </summary>
		private void PS_MM159_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid04 = oForm.Items.Item("Grid04").Specific;
				oGrid04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oGrid05 = oForm.Items.Item("Grid05").Specific;
				oGrid05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

				oForm.DataSources.DataTables.Add("ZTEMP01");
				oForm.DataSources.DataTables.Add("ZTEMP02");
				oForm.DataSources.DataTables.Add("ZTEMP03");
				oForm.DataSources.DataTables.Add("ZTEMP04");
				oForm.DataSources.DataTables.Add("ZTEMP05");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM159_SetComboBox
		/// </summary>
		private void PS_MM159_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
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
		/// PS_MM159_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_MM159_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "작번은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				returnValue = true;
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
			return returnValue;
        }

        ///// <summary>
        ///// PS_MM159_SetGrid
        ///// </summary>
        ///// <param name="GridNo"></param>
        //private void PS_MM159_SetGrid(string GridNo)
        //{
        //    int i;

        //    try
        //    {
        //        oForm.Freeze(true);

        //        switch (GridNo)
        //        {
        //            case "Grid01":
        //                ((SAPbouiCOM.EditTextColumn)oGrid01.Columns.Item(2)).LinkedObjectType = "4"; // Link to ItemMaster
        //                for (i = 0; i <= oGrid01.Columns.Count - 1; i++)
        //                {
        //                    if (oGrid01.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
        //                    {
        //                        oGrid01.Columns.Item(i).RightJustified = true;
        //                    }
        //                }
        //                break;

        //            case "Grid02":
        //                for (i = 0; i <= oGrid02.Columns.Count - 1; i++)
        //                {
        //                    if (oGrid02.DataTable.Columns.Item(i).Type == SAPbouiCOM.BoFieldsType.ft_Float)
        //                    {
        //                        oGrid02.Columns.Item(i).RightJustified = true;
        //                    }
        //                }
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        /// <summary>
        /// PS_MM159_ResizeForm
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void PS_MM159_ResizeForm(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
				//Grid02
				oForm.Items.Item("Grid01").Height = 50;
				oForm.Items.Item("Grid01").Left = 6;
				oForm.Items.Item("Grid01").Width = (oForm.Width / 2) - 15;

				
				//Grid02
				oForm.Items.Item("Grid02").Top = oForm.Items.Item("Grid01").Top + oForm.Items.Item("Grid01").Height + 20;
                oForm.Items.Item("Grid02").Height = (oForm.Height / 2) - 100;
                oForm.Items.Item("Grid02").Left = 6;
                oForm.Items.Item("Grid02").Width = (oForm.Width / 2) - 15;
				
				//Grid02
				oForm.Items.Item("Grid03").Top = oForm.Items.Item("Grid02").Top + oForm.Items.Item("Grid02").Height + 20;
				oForm.Items.Item("Grid03").Height = (oForm.Height / 2) - 100;
				oForm.Items.Item("Grid03").Left = 6;
				oForm.Items.Item("Grid03").Width = (oForm.Width / 2) - 15;

				//Grid04
				oForm.Items.Item("Grid04").Top = oForm.Items.Item("Grid02").Top;
                oForm.Items.Item("Grid04").Height = oForm.Items.Item("Grid02").Height;
                oForm.Items.Item("Grid04").Left = oForm.Items.Item("Grid02").Width + 20;
                oForm.Items.Item("Grid04").Width = (oForm.Width / 2) - 15;


                //Grid05
                oForm.Items.Item("Grid05").Top = oForm.Items.Item("Grid03").Top;
				oForm.Items.Item("Grid05").Height = oForm.Items.Item("Grid03").Height;
                oForm.Items.Item("Grid05").Left = oForm.Items.Item("Grid03").Width + 20;
				oForm.Items.Item("Grid05").Width = (oForm.Width / 2) - 15;


				oForm.Items.Item("static01").Top = oForm.Items.Item("Grid02").Top +10 ;
				oForm.Items.Item("static02").Top = oForm.Items.Item("Grid03").Top + 10;
				oForm.Items.Item("static03").Top = oForm.Items.Item("Grid04").Top + 10;
				oForm.Items.Item("static04").Top = oForm.Items.Item("Grid05").Top + 10;

				oForm.Items.Item("static01").Left = oForm.Items.Item("Grid02").Left;
				oForm.Items.Item("static02").Left = oForm.Items.Item("Grid03").Left;
				oForm.Items.Item("static03").Left = oForm.Items.Item("Grid04").Left;
				oForm.Items.Item("static04").Left = oForm.Items.Item("Grid05").Left;

				if (oGrid01.Rows.Count > 0)
                {
                    oGrid01.AutoResizeColumns();
                }
                if (oGrid02.Rows.Count > 0)
                {
                    oGrid02.AutoResizeColumns();
                }
                if (oGrid03.Rows.Count > 0)
                {
                    oGrid03.AutoResizeColumns();
                }
                if (oGrid04.Rows.Count > 0)
                {
                    oGrid04.AutoResizeColumns();
                }
                if (oGrid05.Rows.Count > 0)
                {
                    oGrid05.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

            /// <summary>
            /// PS_MM159_SearchGrid01Data
            /// </summary>
            private void PS_MM159_Grid01()
		{
			string sQry;
			string ItemCode;  
			try
			{
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				sQry = "SELECT A.CardCode AS 거래처, A.CardName AS 거래처명, CONVERT(CHAR(10),A.DocDate,23) AS 수주일자, B.ItemCode AS 작번,B.Quantity AS 수주수량, B.Price AS 단가, B.Linetotal AS 총계  FROM ORDR A INNER JOIN RDR1 B ON A.DocEntry = B.DocEntry WHERE B.ItemCode ='" + ItemCode + "'";
				oForm.DataSources.DataTables.Item("ZTEMP01").ExecuteQuery(sQry);
				oGrid01.DataTable = oForm.DataSources.DataTables.Item("ZTEMP01");
				oGrid01.AutoResizeColumns();
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
		/// PS_MM159_SearchGrid02Data
		/// </summary>
		private void PS_MM159_Grid02()
		{
			string sQry;
			string ItemCode;
			try
			{
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				sQry = "EXEC [PS_MM158_01] '" + ItemCode + "','1'";
				oForm.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery(sQry);
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("ZTEMP02");
				oGrid02.AutoResizeColumns();
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
		/// PS_MM159_SearchGrid02Data
		/// </summary>
		private void PS_MM159_Grid03()
		{
			string sQry;
			string ItemCode;
			try
			{
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				sQry = "EXEC [PS_MM158_02] '" + ItemCode + "','1'";
				oForm.DataSources.DataTables.Item("ZTEMP03").ExecuteQuery(sQry);
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("ZTEMP03");
				oGrid03.AutoResizeColumns();
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
		/// PS_MM159_SearchGrid02Data
		/// </summary>
		private void PS_MM159_Grid04()
		{
			string sQry;
			string ItemCode;
			try
			{
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				sQry = "EXEC [PS_MM158_01] '" + ItemCode + "','2'";
				oForm.DataSources.DataTables.Item("ZTEMP04").ExecuteQuery(sQry);
				oGrid04.DataTable = oForm.DataSources.DataTables.Item("ZTEMP04");
				oGrid04.AutoResizeColumns();
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
		/// PS_MM159_SearchGrid02Data
		/// </summary>
		private void PS_MM159_Grid05()
		{
			string sQry;
			string ItemCode;
			try
			{
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				sQry = "EXEC [PS_MM158_02] '" + ItemCode + "','2'";
				oForm.DataSources.DataTables.Item("ZTEMP05").ExecuteQuery(sQry);
				oGrid05.DataTable = oForm.DataSources.DataTables.Item("ZTEMP05");
				oGrid05.AutoResizeColumns();
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
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                   // Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "btn01")
					{
						if (PS_MM159_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							if (oGrid02.Rows.Count > 0)
							{
								oGrid02.DataTable.Clear();
							}
							if (oGrid03.Rows.Count > 0)
							{
								oGrid03.DataTable.Clear();
							}
							if (oGrid04.Rows.Count > 0)
							{
								oGrid04.DataTable.Clear();
							}
							if (oGrid05.Rows.Count > 0)
							{
								oGrid05.DataTable.Clear();
							}
						}

						PS_MM159_Grid01();
						PS_MM159_Grid02();
						PS_MM159_Grid03();
						PS_MM159_Grid04();
						PS_MM159_Grid05();
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
					if (pVal.CharPressed == 9)
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처 포맷서치 설정
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //거래처 포맷서치 설정
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    PS_MM159_ResizeForm(FormUID, ref pVal, ref BubbleEvent);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid05);
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
				oForm.Freeze(true);

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
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
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
				oForm.Freeze(false);
			}
		}
	}
}
