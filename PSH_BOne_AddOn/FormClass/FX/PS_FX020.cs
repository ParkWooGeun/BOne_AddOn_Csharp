using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 감가상각계산
	/// </summary>
	internal class PS_FX020 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_FX020H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_FX020L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX020.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_FX020_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_FX020");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				PS_FX020_CreateItems();
				PS_FX020_ComboBox_Setting();

				oForm.EnableMenu("1283", true);	 // 삭제
				oForm.EnableMenu("1287", true);	 // 복제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", true);	 // 행삭제
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
		/// PS_FX020_CreateItems
		/// </summary>
		private void PS_FX020_CreateItems()
		{
			try
			{
				oDS_PS_FX020H = oForm.DataSources.DBDataSources.Item("@PS_FX020H");
				oDS_PS_FX020L = oForm.DataSources.DBDataSources.Item("@PS_FX020L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX020_ComboBox_Setting
		/// </summary>
		private void PS_FX020_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				//자산구분(매트릭스)
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ClasCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "","");
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
		/// PS_FX020_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FX020_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_FX020L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_FX020L.Offset = oRow;
				oDS_PS_FX020L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX020_Copy_MatrixRow
		/// </summary>
		private void PS_FX020_Copy_MatrixRow()
		{
			int i;

			try
			{
				oDS_PS_FX020H.SetValue("Code", 0, "");
				oDS_PS_FX020H.SetValue("Name", 0, "");
				oDS_PS_FX020H.SetValue("U_YM", 0, "");
				oDS_PS_FX020H.SetValue("U_BPLId", 0, "");

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_FX020L.SetValue("Code", i, "");
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX020_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_FX020_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_FX020H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_FX020H.GetValue("U_YM", 0).ToString().Trim()))
                {
					errMessage = "마감년월은 필수입력사항입니다. 확인하세요.";
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
		/// PS_FX020_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_FX020_MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();
				// 라인
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}
				oMat.LoadFromDataSource();
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
		/// PS_FX020_LoadData
		/// </summary>
		private void PS_FX020_LoadData()
		{
			int i;
			string sQry;
			string YM;
			string BPLId;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

				sQry = "EXEC [PS_FX020_01] '" + BPLId + "','" + YM + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_FX020L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_FX020L.Size)
					{
						oDS_PS_FX020L.InsertRecord(i);
					}

					oMat.AddRow();

					oDS_PS_FX020L.Offset = i;
					oDS_PS_FX020L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_FX020L.SetValue("U_ClasCode", i, oRecordSet.Fields.Item("ClasCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_FixCode", i, oRecordSet.Fields.Item("FixCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_SubCode", i, oRecordSet.Fields.Item("SubCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_FixName", i, oRecordSet.Fields.Item("FixName").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_PostDate", i, Convert.ToDateTime(oRecordSet.Fields.Item("PostDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_FX020L.SetValue("U_TeamCode", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_TeamNm", i, oRecordSet.Fields.Item("TeamNm").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_RspCode", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_RspNm", i, oRecordSet.Fields.Item("RspNm").Value.ToString().Trim());

					oDS_PS_FX020L.SetValue("U_PrcCode", i, oRecordSet.Fields.Item("PrcCode").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_PrcName", i, oRecordSet.Fields.Item("PrcName").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_LongYear", i, oRecordSet.Fields.Item("LongYear").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_DepRate", i, oRecordSet.Fields.Item("DepRate").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_PostAmt", i, oRecordSet.Fields.Item("PostAmt").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_OBalance", i, oRecordSet.Fields.Item("OBalance").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_HisAmt", i, oRecordSet.Fields.Item("HisAmt").Value.ToString().Trim());

					oDS_PS_FX020L.SetValue("U_GatAmt", i, oRecordSet.Fields.Item("GatAmt").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_EBalance", i, oRecordSet.Fields.Item("EBalance").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_FixYAmt", i, oRecordSet.Fields.Item("FixYAmt").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_FixAmt", i, oRecordSet.Fields.Item("FixAmt").Value.ToString().Trim());
					oDS_PS_FX020L.SetValue("U_FixMAmt", i, oRecordSet.Fields.Item("FixMAmt").Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				ProgressBar01.Stop();
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string BPLId;
			string YM;
			string Code;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_FX020_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_FX020_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							YM = oDS_PS_FX020H.GetValue("U_YM", 0).ToString().Trim();
							BPLId = oDS_PS_FX020H.GetValue("U_BPLId", 0).ToString().Trim();
							Code = YM + BPLId;
							oDS_PS_FX020H.SetValue("Code", 0, Code);
							oDS_PS_FX020H.SetValue("Name", 0, Code);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
						}
					}
					else if (pVal.ItemUID == "Btn01")
					{
						if (PS_FX020_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}

						BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
						YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
						sQry = "Select Cnt = Count(*) From [@PS_FX025H] Where U_BPLId = '" + BPLId + "' And U_YM = '" + YM + "' And Isnull(U_jdtCC,'N') = 'Y'";
						oRecordSet.DoQuery(sQry);

						if (oRecordSet.Fields.Item("Cnt").Value <= 0)
						{
							PS_FX020_LoadData();
						}
						else
						{
							PSH_Globals.SBO_Application.MessageBox("분개처리되어 작업을 할 수 없습니다.");
							BubbleEvent = false;
							return;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					if (pVal.Row == 0)
					{
						//정렬
						oMat.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
						oMat.FlushToDataSource();
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
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_FX020_Add_MatrixRow(oMat.RowCount, false);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX020H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX020L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;
			string BPLId;
			string YM;
			string Code;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
					YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
					sQry = "Select Cnt = Count(*) From [@PS_FX025H] Where U_BPLId = '" + BPLId + "' And U_YM = '" + YM + "' And Isnull(U_jdtCC,'N') = 'Y'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.Fields.Item("Cnt").Value <= 0)
					{
						Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
						sQry = "DELETE FROM Z_PS_FX020L WHERE Code = '" + Code + "'";
						oRecordSet.DoQuery(sQry);
					}
					else
					{
						PSH_Globals.SBO_Application.MessageBox("분개처리된 자료는 삭제 할 수 없습니다.");
						BubbleEvent = false;
						return;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (oMat.RowCount != oMat.VisualRowCount)
					{
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
						}

						oMat.FlushToDataSource();
						oDS_PS_FX020L.RemoveRecord(oDS_PS_FX020L.Size - 1);
						oMat.Clear();
						oMat.LoadFromDataSource();

						if (!string.IsNullOrEmpty(oMat.Columns.Item("FixCode").Cells.Item(oMat.RowCount).Specific.Value.ToString().Trim()))
						{
							PS_FX020_Add_MatrixRow(oMat.RowCount, false);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// Raise_MenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_MenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							PS_FX020_Add_MatrixRow(oMat.VisualRowCount, false);
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							break;	 
									 
						case "1284": //취소
							break;	 
						case "1286": //닫기
							break;	 
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							break;	 
						case "1282": //추가
							PS_FX020_Add_MatrixRow(0, true);
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1287": // 복제
							PS_FX020_Copy_MatrixRow();
							break;
						case "7169": //엑셀 내보내기
							oDS_PS_FX020L.RemoveRecord(oDS_PS_FX020L.Size - 1);
							oMat.LoadFromDataSource();
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
		/// Raise_FormDataEvent
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
