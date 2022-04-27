using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// MG 계량치 분석값 등록
	/// </summary>
	internal class PS_QM320 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM320H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM320L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM320.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM320_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM320");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM320_CreateItems();
				PS_QM320_ComboBox_Setting();
				PS_QM320_FormReset();
				PS_QM320_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_QM320_CreateItems
		/// </summary>
		private void PS_QM320_CreateItems()
		{
			try
			{
				oDS_PS_QM320H = oForm.DataSources.DBDataSources.Item("@PS_QM320H");
				oDS_PS_QM320L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");
				//사번
				oForm.DataSources.UserDataSources.Add("SMachCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SMachCod").Specific.DataBind.SetBound(true, "", "SMachCod");
				//지급일자(시작)
				oForm.DataSources.UserDataSources.Add("SPrvdDtFr", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SPrvdDtFr").Specific.DataBind.SetBound(true, "", "SPrvdDtFr");
				oForm.Items.Item("SPrvdDtFr").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				//지급일자(종료)
				oForm.DataSources.UserDataSources.Add("SPrvdDtTo", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SPrvdDtTo").Specific.DataBind.SetBound(true, "", "SPrvdDtTo");
				oForm.Items.Item("SPrvdDtTo").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM320_ComboBox_Setting
		/// </summary>
		private void PS_QM320_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//기본정보-사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("MachCode").Specific.ValidValues.Add("122201", "연속소둔로");
				oForm.Items.Item("MachCode").Specific.ValidValues.Add("122901", "S&D 1호기");
				oForm.Items.Item("MachCode").Specific.ValidValues.Add("122902", "S&D 2호기");
				oForm.Items.Item("MachCode").Specific.ValidValues.Add("122903", "S&D 3호기");
				oForm.Items.Item("MachCode").Specific.ValidValues.Add("122604", "DEGREASER 4호기");
				oForm.Items.Item("MachCode").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("SMachCod").Specific.ValidValues.Add("122201", "연속소둔로");
				oForm.Items.Item("SMachCod").Specific.ValidValues.Add("122901", "S&D 1호기");
				oForm.Items.Item("SMachCod").Specific.ValidValues.Add("122902", "S&D 2호기");
				oForm.Items.Item("SMachCod").Specific.ValidValues.Add("122903", "S&D 3호기");
				oForm.Items.Item("SMachCod").Specific.ValidValues.Add("122604", "DEGREASER 4호기");
				oForm.Items.Item("SMachCod").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//조회조건-사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		///  화면 초기화
		/// </summary>
		private void PS_QM320_FormReset()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				//관리번호
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_QM320H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oDS_PS_QM320H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_QM320H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
				}

				oDS_PS_QM320H.SetValue("U_LotId", 0, "");	//수량
				oDS_PS_QM320H.SetValue("U_HValue", 0, "0"); //반납일자
				oDS_PS_QM320H.SetValue("U_DValue", 0, "0"); //반납사유
				oDS_PS_QM320H.SetValue("U_NValue", 0, "0"); //비고
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_QM320_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
					oForm.Items.Item("BtnDelete").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
					oForm.Items.Item("BtnDelete").Enabled = true;
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
		/// SD 공정일 경우 DValue, NValue/ 열처리 공정일 경우 LOTID, HValue 입력되도록 함.
		/// </summary>
		private void PS_QM320_DataCheck()
		{
			string MachCode;

			try
			{
				MachCode = oForm.Items.Item("MachCode").Specific.Value.ToString().Trim();

				if (MachCode == "122201")
				{
					oForm.Items.Item("DValue").Visible = false;
					oForm.Items.Item("NValue").Visible = false;
					oForm.Items.Item("LotId").Visible = true;
					oForm.Items.Item("HValue").Visible = true;
				}
				else
				{
					oForm.Items.Item("DValue").Visible = true;
					oForm.Items.Item("NValue").Visible = true;
					oForm.Items.Item("LotId").Visible = false;
					oForm.Items.Item("HValue").Visible = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 데이터 조회
		/// </summary>
		private void PS_QM320_MTX01()
		{
			int i;
			string SBPLID;    //사업장
			string SMachCod;  //장비코드
			string SPrvdDtFr; //일자(시작)
			string SPrvdDtTo; //일자(종료)
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
				SMachCod = oForm.Items.Item("SMachCod").Specific.Value.ToString().Trim();
				SPrvdDtFr = oForm.Items.Item("SPrvdDtFr").Specific.Value.ToString().Trim();
				SPrvdDtTo = oForm.Items.Item("SPrvdDtTo").Specific.Value.ToString().Trim();

				if (SBPLID == "%")
				{
					SBPLID = "";
				}

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_QM320_01] '" + SBPLID + "','" + SMachCod + "','" + SPrvdDtFr + "','" + SPrvdDtTo + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM320L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_QM320_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM320L.Size)
					{
						oDS_PS_QM320L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_QM320L.Offset = i;

					oDS_PS_QM320L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM320L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
					oDS_PS_QM320L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());    //사업장
					oDS_PS_QM320L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("MachCode").Value.ToString().Trim()); //장비코드
					oDS_PS_QM320L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //일자
					oDS_PS_QM320L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("LotId").Value.ToString().Trim());  //로트번호
					oDS_PS_QM320L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("HValue").Value.ToString().Trim()); //열처리 측정값
					oDS_PS_QM320L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("DValue").Value.ToString().Trim()); //주간측정값
					oDS_PS_QM320L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("NValue").Value.ToString().Trim()); //야간 측정값
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 기본정보 삭제
		/// </summary>
		private void PS_QM320_DeleteData()
		{
			string DocEntry;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

					sQry = "SELECT COUNT(*) FROM [@PS_QM320H] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "DELETE FROM [@PS_QM320H] WHERE DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// 기본정보를 수정
		/// </summary>
		/// <returns></returns>
		private bool PS_QM320_UpdateData()
		{
			bool ReturnValue = false;
			int DocEntry;
			string LotId;  //사업장
			string HValue; //사번
			string DValue; //성명
			string NValue; //팀
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
				LotId = oForm.Items.Item("LotId").Specific.Value.ToString().Trim();
				HValue = oForm.Items.Item("HValue").Specific.Value.ToString().Trim();
				DValue = oForm.Items.Item("DValue").Specific.Value.ToString().Trim();
				NValue = oForm.Items.Item("NValue").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(Convert.ToString(DocEntry)))
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요.";
					throw new Exception();
				}

				sQry = "    UPDATE   [@PS_QM320H]";
				sQry += " SET         U_LotId = '" + LotId + "',";
				sQry += "              U_HValue = '" + HValue + "',";
				sQry += "              U_DValue = '" + DValue + "',";
				sQry += "              U_NValue = '" + NValue + "'";
				sQry += " WHERE    DocEntry = '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

				ReturnValue = true;
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_QM320_AddData()
		{
			bool ReturnValue = false;
			int DocEntry;
			string BPLId;    //사업장
			string DocDate;	 //사번
			string MachCode; //성명
			string LotId;    //팀
			string HValue;   //팀명
			string DValue;   //담당
			string NValue;   //담당명
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				MachCode = oForm.Items.Item("MachCode").Specific.Value.ToString().Trim();
				LotId = oForm.Items.Item("LotId").Specific.Value.ToString().Trim();
				HValue = oForm.Items.Item("HValue").Specific.Value.ToString().Trim();
				DValue = oForm.Items.Item("DValue").Specific.Value.ToString().Trim();
				NValue = oForm.Items.Item("NValue").Specific.Value.ToString().Trim();

				sQry = "select * from [@PS_QM320H] where u_DocDate = '" + DocDate + "' and " + "u_MachCode ='" + MachCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount >= 1)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "중복 입력입니다. 확인하세요.";
					throw new Exception();
				}
				else
				{
					//DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
					sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_QM320H]";
					oRecordSet.DoQuery(sQry);

					if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
					{
						DocEntry = 1;
					}
					else
					{
						DocEntry = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
					}

					sQry = " INSERT INTO [@PS_QM320H]";
					sQry += " (";
					sQry += "     DocEntry,";
					sQry += "     DocNum,";
					sQry += "     U_BPLId,";
					sQry += "     U_DocDate,";
					sQry += "     U_MachCode,";
					sQry += "     U_LotId,";
					sQry += "     U_HValue,";
					sQry += "     U_DValue,";
					sQry += "     U_NValue";
					sQry += " )";
					sQry += " VALUES";
					sQry += " (";
					sQry += DocEntry + ",";
					sQry += DocEntry + ",";
					sQry += "'" + BPLId + "',";
					sQry += "'" + DocDate + "',";
					sQry += "'" + MachCode + "',";
					sQry += "'" + LotId + "',";
					sQry += "'" + HValue + "',";
					sQry += "'" + DValue + "',";
					sQry += "'" + NValue + "'";
					sQry += ")";
					oRecordSet.DoQuery(sQry);

					PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				}

				ReturnValue = true;
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
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
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM320_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM320_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM320_LoadCaption();
							PS_QM320_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM320_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM320_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM320_LoadCaption();
							PS_QM320_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						PS_QM320_FormReset();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_QM320_LoadCaption();
						PS_QM320_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_QM320_DeleteData();
							PS_QM320_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM320_LoadCaption();
							PS_QM320_MTX01();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "PrtrCd", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SCntcCode", "");
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "MachCode")
					{
						PS_QM320_DataCheck();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
							oDS_PS_QM320H.SetValue("DocEntry", 0, oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_BPLId", 0, oMat.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_MachCode", 0, oMat.Columns.Item("MachCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_DocDate", 0, oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_LotId", 0, oMat.Columns.Item("LotId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_HValue", 0, oMat.Columns.Item("HValue").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_DValue", 0, oMat.Columns.Item("DValue").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_QM320H.SetValue("U_NValue", 0, oMat.Columns.Item("NValue").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_QM320_LoadCaption();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					PS_QM320_DataCheck();
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
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM320H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM320L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
						case "1283": //제거
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
							PS_QM320_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_QM320_LoadCaption();
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
							break;
						case "1287": //복제
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
