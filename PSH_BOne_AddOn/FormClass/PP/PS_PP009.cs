using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	///  공정인원등록
	/// </summary>
	internal class PS_PP009 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP009L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP009.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP009_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP009");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP009_CreateItems();
				PS_PP009_ComboBox_Setting();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_PP009_CreateItems
		/// </summary>
		private void PS_PP009_CreateItems()
		{
			try
			{
				oDS_PS_PP009L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//팀
				oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

				//담당
				oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

				//반
				oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

				//사번
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

				//직원구분
				oForm.DataSources.UserDataSources.Add("EmpType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("EmpType").Specific.DataBind.SetBound(true, "", "EmpType");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP009_ComboBox_Setting
		/// </summary>
		private void PS_PP009_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//직원구분
				oForm.Items.Item("EmpType").Specific.ValidValues.Add("%", "선택");
				sQry = " SELECT    U_Code,";
				sQry += "           U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = 'P126'";
				sQry += "           AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("EmpType").Specific, sQry, "", false, false);
				oForm.Items.Item("EmpType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//매트릭스
				//사업장
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");

				//팀
				sQry = " SELECT    U_Code,";
				sQry += "           U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '1'";
				sQry += "           AND U_UseYN = 'Y'";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("TeamCode"), sQry, "", "");

				//담당
				sQry = " SELECT    U_Code,";
				sQry += "           U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '2'";
				sQry += "           AND U_UseYN = 'Y'";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("RspCode"), sQry, "", "");

				//반
				sQry = " SELECT    U_Code,";
				sQry += "           U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '9'";
				sQry += "           AND U_UseYN = 'Y'";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ClsCode"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP009_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		public void PS_PP009_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP009L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP009L.Offset = oRow;
				oDS_PS_PP009L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP009_MTX01
		/// 데이터 조회
		/// </summary>
		private void PS_PP009_MTX01()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string BPLID;    //사업장
			string TeamCode; //팀
			string RspCode;  //담당
			string ClsCode;  //반
			string CntcCode; //사번
			string EmpType;  //직원구분
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				EmpType = oForm.Items.Item("EmpType").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "     EXEC [PS_PP009_01]";
				sQry += "'" + BPLID + "',";
				sQry += "'" + TeamCode + "',";
				sQry += "'" + RspCode + "',";
				sQry += "'" + ClsCode + "',";
				sQry += "'" + CntcCode + "',";
				sQry += "'" + EmpType + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP009L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					PS_PP009_Add_MatrixRow(0, true);
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP009L.Size)
					{
						oDS_PS_PP009L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP009L.Offset = i;

					oDS_PS_PP009L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP009L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());    //사업장
					oDS_PS_PP009L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim()); //팀
					oDS_PS_PP009L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());  //담당
					oDS_PS_PP009L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ClsCode").Value.ToString().Trim());  //반
					oDS_PS_PP009L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim()); //사번
					oDS_PS_PP009L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim()); //성명
					oDS_PS_PP009L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("MCpCode").Value.ToString().Trim());  //주공정코드
					oDS_PS_PP009L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("MCpName").Value.ToString().Trim());  //주공정명
					oDS_PS_PP009L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("SCpCode").Value.ToString().Trim());  //부공정코드
					oDS_PS_PP009L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("SCpName").Value.ToString().Trim());  //부공정명

					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP009_UpdateData
		/// 데이터 입력 및 수정
		/// </summary>
		/// <returns></returns>
		private bool PS_PP009_UpdateData()
		{
			bool functionReturnValue = false;

			int loopCount;
			string sQry;

			string BPLID;	 //사업장
			string TeamCode; //팀
			string RspCode;	 //담당
			string ClsCode;	 //반
			string CntcCode; //사번
			string CntcName; //성명
			string MCpCode;	 //주공정코드
			string MCpName;	 //주공정명
			string SCpCode;	 //부공정코드
			string SCpName;  //부공정명

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				for (loopCount = 1; loopCount <= oMat.RowCount; loopCount++)
				{
					BPLID = oMat.Columns.Item("BPLId").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					TeamCode = oMat.Columns.Item("TeamCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					RspCode = oMat.Columns.Item("RspCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					ClsCode = oMat.Columns.Item("ClsCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					CntcCode = oMat.Columns.Item("CntcCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					CntcName = oMat.Columns.Item("CntcName").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					MCpCode = oMat.Columns.Item("MCpCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					MCpName = oMat.Columns.Item("MCpName").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					SCpCode = oMat.Columns.Item("SCpCode").Cells.Item(loopCount).Specific.Value.ToString().Trim();
					SCpName = oMat.Columns.Item("SCpName").Cells.Item(loopCount).Specific.Value.ToString().Trim();

					sQry = "     EXEC [PS_PP009_02]";
					sQry += "'" + BPLID + "',";
					sQry += "'" + TeamCode + "',";
					sQry += "'" + RspCode + "',";
					sQry += "'" + ClsCode + "',";
					sQry += "'" + CntcCode + "',";
					sQry += "'" + CntcName + "',";
					sQry += "'" + MCpCode + "',";
					sQry += "'" + MCpName + "',";
					sQry += "'" + SCpCode + "',";
					sQry += "'" + SCpName + "'";
					oRecordSet.DoQuery(sQry);

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oMat.RowCount + "건 저장 중...!";
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
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
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnModify")
					{
						if (PS_PP009_UpdateData() == false)
						{
							BubbleEvent = false;
							return;
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_PP009_MTX01();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MCpCode");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SCpCode");
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
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int loopCount;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						switch (pVal.ItemUID)
						{
							case "BPLId":
								if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
								{
									for (loopCount = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
									{
										oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
									}
								}

								oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "선택");
								sQry = " SELECT    U_Code,";
								sQry += "           U_CodeNm";
								sQry += " FROM      [@PS_HR200L]";
								sQry += " WHERE     Code = '1'";
								sQry += "           AND U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_UseYN = 'Y'";
								sQry += " ORDER BY  U_Seq";
								dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
								oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
								oForm.Items.Item("TeamCode").DisplayDesc = true;
								break;

							case "TeamCode":
								if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
								{
									for (loopCount = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
									{
										oForm.Items.Item("RspCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
									}
								}

								oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "선택");
								sQry = " SELECT    U_Code,";
								sQry += "           U_CodeNm";
								sQry += " FROM      [@PS_HR200L]";
								sQry += " WHERE     Code = '2'";
								sQry += "           AND U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_UseYN = 'Y'";
								sQry += " ORDER BY  U_Seq";
								dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
								oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
								oForm.Items.Item("RspCode").DisplayDesc = true;
								break;

							case "RspCode":
								if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
								{
									for (loopCount = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
									{
										oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
									}
								}

								oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "선택");
								sQry = " SELECT    U_Code,";
								sQry += "           U_CodeNm";
								sQry += " FROM      [@PS_HR200L]";
								sQry += " WHERE     Code = '9'";
								sQry += "           AND U_Char3 = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_Char2 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
								sQry += "           AND U_UseYN = 'Y'";
								sQry += " ORDER BY  U_Seq";
								dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
								oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
								oForm.Items.Item("ClsCode").DisplayDesc = true;
								break;
						}
					}
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "MCpCode")
							{
								oDS_PS_PP009L.SetValue("U_ColReg07", pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_PP009L.SetValue("U_ColReg08", pVal.Row - 1, dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oMat.Columns.Item("MCpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));

							} else if (pVal.ColUID == "SCpCode")
							{
								oDS_PS_PP009L.SetValue("U_ColReg09", pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_PP009L.SetValue("U_ColReg10", pVal.Row - 1, dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oMat.Columns.Item("SCpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
							}
						}
						else
						{
						}
						oMat.LoadFromDataSource();
						oMat.AutoResizeColumns();
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
			finally
            {
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else
				{
					oLastItemUID01 = pVal.ItemUID;
					oLastColUID01 = "";
					oLastColRow01 = 0;
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP009L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = pVal.ColUID;
						oLastColRow01 = pVal.Row;
					}
				}
				else
				{
					oLastItemUID01 = pVal.ItemUID;
					oLastColUID01 = "";
					oLastColRow01 = 0;
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
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							break;
						case "1285": //복원
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
							PS_PP009_Add_MatrixRow(oMat.VisualRowCount, false);
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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
							oDS_PS_PP009L.RemoveRecord(oDS_PS_PP009L.Size - 1);
							oMat.LoadFromDataSource();
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
