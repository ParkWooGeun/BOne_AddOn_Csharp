using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검사성적서출력(신양식)
	/// </summary>
	internal class PS_QM041 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_QM041L; //등록라인

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oYM01"></param>
		/// <param name="oFormDocEntry01"></param>
		public void LoadForm(string oYM01, string oFormDocEntry01)
		{
			this.MainLoadForm(oYM01, oFormDocEntry01);
		}

        public override void LoadForm(string oFormDocEntry01)
        {
			this.MainLoadForm("", oFormDocEntry01);
        }

		private void MainLoadForm(string oYM01, string oFormDocEntry01)
        {
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM041.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM041_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM041");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();

				oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
				oForm.Items.Item("YYYYMM").Specific.Value = oYM01;

				if (!string.IsNullOrEmpty(oFormDocEntry01))
				{
					oForm.Items.Item("Gubun").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
				}
				else
				{
					oForm.Items.Item("Gubun").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
				}

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1284"), false); // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제

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
        /// CreateItems
        /// </summary>
        private void CreateItems()
		{
			try
			{
				oDS_PS_QM041L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");


				oForm.DataSources.UserDataSources.Add("YYYYMM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("YYYYMM").Specific.DataBind.SetBound(true, "", "YYYYMM");
				oForm.Items.Item("YYYYMM").Specific.Value = DateTime.Now.ToString("yyyy-MM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] Where BPLId in ('1', '2') order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//기본사업장SETTING
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "Packing기준");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("2", "납품기준");
				oForm.Items.Item("Gubun").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
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
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim()))
				{
					errMessage = "조회년월은 필수입니다. 입력하여 주십시오.";
					throw new Exception();
				}
				if (oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim().Length != 7)
				{
					errMessage = "조회년월의 자리수(YYYY-MM)를 확인하여 주십시오.";
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Search_Matrix_Data
		/// </summary>
		private void Search_Matrix_Data()
		{
			string sQry;

			int j;
			int Cnt;
			string BPLId;
			string YYYYMM;
			string Gubun;
			string DocEntry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
				Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

				sQry = "EXEC PS_QM041_01 '" + BPLId + "', '" + YYYYMM + "', '" + Gubun + "', '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				Cnt = oDS_PS_QM041L.Size;
				if (Cnt > 0)
				{
					for (j = 0; j <= Cnt - 1; j++)
					{
						oDS_PS_QM041L.RemoveRecord(oDS_PS_QM041L.Size - 1);
					}
					if (Cnt == 1)
					{
						oDS_PS_QM041L.Clear();
					}
				}
				oMat.LoadFromDataSource();

				j = 1;
				while (!(oRecordSet.EoF))
				{
					if (oDS_PS_QM041L.Size < j)
					{
						oDS_PS_QM041L.InsertRecord(j - 1); //라인추가
					}
					oDS_PS_QM041L.SetValue("U_LineNum", j - 1, Convert.ToString(j));
					if (Gubun == "2")
					{
						oDS_PS_QM041L.SetValue("U_ColReg01", j - 1, "Y");
					}
					else
					{
						oDS_PS_QM041L.SetValue("U_ColReg01", j - 1, "N");
					}
					oDS_PS_QM041L.SetValue("U_ColReg02", j - 1, oRecordSet.Fields.Item("U_PackNo").Value.ToString().Trim());
					oDS_PS_QM041L.SetValue("U_ColReg03", j - 1, oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim());
					oDS_PS_QM041L.SetValue("U_ColReg04", j - 1, oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim());
					oDS_PS_QM041L.SetValue("U_ColReg05", j - 1, oRecordSet.Fields.Item("U_CardCode").Value.ToString().Trim());
					oDS_PS_QM041L.SetValue("U_ColReg06", j - 1, oRecordSet.Fields.Item("U_CardName").Value.ToString().Trim());
					oDS_PS_QM041L.SetValue("U_ColReg07", j - 1, oRecordSet.Fields.Item("Type").Value.ToString().Trim());
					j += 1;
					oRecordSet.MoveNext();
				}
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Print_Query
		/// </summary>
		[STAThread]
		private void Print_Query()
		{
			int i;
			string WinTitle;
			string ReportName;
			string sQry;

			string BPLId;
			string Chk;
			string ItemCode = string.Empty;
			string CardCode = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				// 임시테이블에 check된항목저장
				sQry = "Delete [Z_PS_QM040] WHERE BPLId = '" + BPLId + "'";
				oRecordSet.DoQuery(sQry);

				oMat.FlushToDataSource();
				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (oDS_PS_QM041L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						sQry = "Insert [Z_PS_QM040] values ('" + BPLId + "', '" + oDS_PS_QM041L.GetValue("U_ColReg02", i).ToString().Trim() + "')";
						oRecordSet.DoQuery(sQry);

						CardCode = oDS_PS_QM041L.GetValue("U_ColReg05", i).ToString().Trim(); // (주)TSD '12440' 찿기위해 MOVE
						ItemCode = oDS_PS_QM041L.GetValue("U_ColReg03", i).ToString().Trim();
					}
				}

				// B/G타입  체크
				if (Convert.ToInt32(dataHelpClass.GetValue("SELECT count(*) FROM [@PS_PP090H] a inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry INNER JOIN [Z_PS_QM040] z on a.U_BPLId = z.BPLId and a.U_PackNo = z.PackNo left  join [OITM] c on b.U_ItemCode = c.ItemCode WHERE z.BPLId = '" + BPLId + "' and c.U_ItemType in ('16','17','19')", 0, 0)) > 0)
				{
					Chk = "Y";
				}
				else
				{
					Chk = "N";
				}

				WinTitle = "[PS_QM041] 검사성적서출력(신)";

				if (Chk == "Y")
				{
					ReportName = "PS_QM041_02.RPT"; // B/G 타입
				}
				else
				{
					if (CardCode != "12440")
					{
						ReportName = "PS_QM041_01.RPT"; // 일반
					}
					else
					{
						if (ItemCode != "104010098")
						{
							ReportName = "PS_QM041_03.RPT"; // (주)TSD
						}
						else
						{
							ReportName = "PS_QM041_04.RPT"; // (주)TSD "104010098"
						}
					}
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));

				//SubReport Parameter
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLId", BPLId, "PS_QM040_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLId", BPLId, "PS_QM040_SUB_02"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLId", BPLId, "PS_QM040_SUB_03"));

				formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
					}
					else if (pVal.ItemUID == "Search")
					{
						if (HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							Search_Matrix_Data();
						}
					}
					else if (pVal.ItemUID == "Print")
					{
						System.Threading.Thread thread = new System.Threading.Thread(Print_Query);
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							if (pVal.ColUID == "Check")
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									if (oDS_PS_QM041L.GetValue("U_ColReg01", i).ToString().Trim() == "N")
									{
										oDS_PS_QM041L.SetValue("U_ColReg01", i, "Y");
									}
									else
									{
										oDS_PS_QM041L.SetValue("U_ColReg01", i, "N");
									}
								}
								oMat.LoadFromDataSource();
							}
							else
							{
								oMat.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
								oMat.FlushToDataSource();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM041L);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
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
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							oForm.Items.Item("DocEntry").Specific.VALUE = "";
							oForm.Items.Item("Gubun").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("YYYYMM").Specific.VALUE = DateTime.Now.ToString("yyyy-MM");
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
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
				if (BusinessObjectInfo.BeforeAction == true)
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
