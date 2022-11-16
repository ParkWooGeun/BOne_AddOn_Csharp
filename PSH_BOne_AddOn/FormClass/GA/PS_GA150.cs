using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 안전보호구지급 이력 등록
	/// </summary>
	internal class PS_GA150 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_GA150H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_GA150L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_GA150.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_GA150_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_GA150");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_GA150_CreateItems();
				PS_GA150_ComboBox_Setting();
				PS_GA150_FormReset();
				PS_GA150_LoadCaption();

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
		/// PS_GA150_CreateItems
		/// </summary>
		private void PS_GA150_CreateItems()
		{
			try
			{
				oDS_PS_GA150H = oForm.DataSources.DBDataSources.Item("@PS_GA150H");
				oDS_PS_GA150L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");

				//사번
				oForm.DataSources.UserDataSources.Add("SCntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SCntcCode").Specific.DataBind.SetBound(true, "", "SCntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("SCntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("SCntcName").Specific.DataBind.SetBound(true, "", "SCntcName");

				//지급일자(시작)
				oForm.DataSources.UserDataSources.Add("SPrvdDtFr", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SPrvdDtFr").Specific.DataBind.SetBound(true, "", "SPrvdDtFr");

				//지급일자(종료)
				oForm.DataSources.UserDataSources.Add("SPrvdDtTo", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SPrvdDtTo").Specific.DataBind.SetBound(true, "", "SPrvdDtTo");

				//등록일
				oForm.DataSources.UserDataSources.Add("SRegDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SRegDt").Specific.DataBind.SetBound(true, "", "SRegDt");

				//최근등록자료 조회
				oForm.DataSources.UserDataSources.Add("ChkMaxDt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("ChkMaxDt").Specific.DataBind.SetBound(true, "", "ChkMaxDt");
				oForm.Items.Item("ChkMaxDt").Specific.Checked = false;

				//일자 SAET
				oForm.Items.Item("PrvdDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("RegDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("SRegDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("CntcCode").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_GA150_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//기본정보-사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//조회조건-사업장
				oForm.Items.Item("SBPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//매트릭스-사업장
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 화면 초기화
		/// </summary>
		private void PS_GA150_FormReset()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//관리번호
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA150H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oDS_PS_GA150H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_GA150H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
				}

				oDS_PS_GA150H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID()); //사업장
				oDS_PS_GA150H.SetValue("U_CntcCode", 0, ""); //사번
				oDS_PS_GA150H.SetValue("U_CntcName", 0, ""); //성명
				oDS_PS_GA150H.SetValue("U_TeamCd", 0, "");   //팀
				oDS_PS_GA150H.SetValue("U_TeamNm", 0, "");   //팀명
				oDS_PS_GA150H.SetValue("U_RspCd", 0, "");    //담당
				oDS_PS_GA150H.SetValue("U_RspNm", 0, "");    //담당명
				oDS_PS_GA150H.SetValue("U_PrtrCd", 0, "");   //보호구코드
				oDS_PS_GA150H.SetValue("U_PrtrNm", 0, "");   //보호구명
				oDS_PS_GA150H.SetValue("U_PrtrSpec", 0, ""); //보호구규격
				oDS_PS_GA150H.SetValue("U_PrvdDt", 0, DateTime.Now.ToString("yyyyMMdd")); //지급일자
				oDS_PS_GA150H.SetValue("U_LPrvdDt", 0, ""); //전지급일
				oDS_PS_GA150H.SetValue("U_Pred", 0, "");    //주기
				oDS_PS_GA150H.SetValue("U_Qty", 0, "1");    //수량
				oDS_PS_GA150H.SetValue("U_RtrnDt", 0, "");  //반납일자
				oDS_PS_GA150H.SetValue("U_RtrnRsn", 0, ""); //반납사유
				oDS_PS_GA150H.SetValue("U_Note", 0, "");    //비고
				oDS_PS_GA150H.SetValue("U_RegDt", 0, DateTime.Now.ToString("yyyyMMdd")); //등록일
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
		private void PS_GA150_LoadCaption()
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
		/// 중복데이터 체크
		/// </summary>
		/// <returns></returns>
		private bool PS_GA150_DataCheck()
		{
			bool ReturnValue = false;
			string BPLID;
			string CntcCode;
			string PrvdDt;
			string ItemCode;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				PrvdDt = oForm.Items.Item("PrvdDt").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim();

				sQry = " EXEC PS_GA150_02 '";
				sQry += BPLID + "','";
				sQry += CntcCode + "','";
				sQry += PrvdDt + "','";
				sQry += ItemCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item("ReturnValue").Value.ToString().Trim() == "True")
				{
					ReturnValue = true;
				}
				else
				{
					errMessage = "이미입력된 데이터입니다. 확인하세요.";
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
		/// 데이터 조회
		/// </summary>
		private void PS_GA150_MTX01()
		{
			int i;
			string SBPLID;    //사업장
			string SCntcCode; //사번
			string SPrvdDtFr; //지급일자(시작)
			string SPrvdDtTo; //지급일자(종료)
			string SRegDt;    //등록일
			string ChkMaxDt;  //최근등록자료 조회
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				if (oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim() == "%")
				{
					SBPLID = "";
				}
				else
				{
					SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
				}
				SCntcCode = oForm.Items.Item("SCntcCode").Specific.Value.ToString().Trim();
				SPrvdDtFr = oForm.Items.Item("SPrvdDtFr").Specific.Value.ToString().Trim();
				SPrvdDtTo = oForm.Items.Item("SPrvdDtTo").Specific.Value.ToString().Trim();
				SRegDt = oForm.Items.Item("SRegDt").Specific.Value.ToString().Trim();

				//최근등록자료 조회
				if (oForm.DataSources.UserDataSources.Item("ChkMaxDt").Value.ToString().Trim() == "Y")
				{
					ChkMaxDt = "Y";
				}
				else
				{
					ChkMaxDt = "N";
				}

				ProgressBar01.Text = "조회시작!";

				if (ChkMaxDt == "N")
				{
					sQry = " EXEC [PS_GA150_01] '";
					sQry += SBPLID + "','";
					sQry += SCntcCode + "','";
					sQry += SPrvdDtFr + "','";
					sQry += SPrvdDtTo + "','";
					sQry += SRegDt + "'";
				}
				else
				{
					sQry = " EXEC [PS_GA150_03] '";
					sQry += SBPLID + "','";
					sQry += SCntcCode + "','";
					sQry += SPrvdDtFr + "','";
					sQry += SPrvdDtTo + "','";
					sQry += SRegDt + "'";
				}
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_GA150L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_GA150_LoadCaption();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_GA150L.Size)
					{
						oDS_PS_GA150L.InsertRecord(i);
					}
					oMat.AddRow();
					oDS_PS_GA150L.Offset = i;
					oDS_PS_GA150L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_GA150L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
					oDS_PS_GA150L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());    //사업장
					oDS_PS_GA150L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim()); //사번
					oDS_PS_GA150L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim()); //성명
					oDS_PS_GA150L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("TeamCd").Value.ToString().Trim());   //팀
					oDS_PS_GA150L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("TeamNm").Value.ToString().Trim());   //팀명
					oDS_PS_GA150L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("RspCd").Value.ToString().Trim());    //담당
					oDS_PS_GA150L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("RspNm").Value.ToString().Trim());    //담당명
					oDS_PS_GA150L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("PrtrCd").Value.ToString().Trim());   //보호구코드
					oDS_PS_GA150L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("PrtrNm").Value.ToString().Trim());   //보호구명
					oDS_PS_GA150L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("PrtrSpec").Value.ToString().Trim()); //보호구규격
					oDS_PS_GA150L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("PrvdDt").Value.ToString().Trim()).ToString("yyyyMMdd"));  //지급일자
					oDS_PS_GA150L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("LPrvdDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //전지급일
					oDS_PS_GA150L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Pred").Value.ToString().Trim());    //주기
					oDS_PS_GA150L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Qty").Value.ToString().Trim());     //수량
					oDS_PS_GA150L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("RtrnDt").Value.ToString().Trim());  //반납일자
					oDS_PS_GA150L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("RtrnRsn").Value.ToString().Trim()); //반납사유
					oDS_PS_GA150L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Note").Value.ToString().Trim());    //비고
					oDS_PS_GA150L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("RegDt").Value.ToString().Trim());   //등록일
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
		private void PS_GA150_DeleteData()
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

					sQry = "SELECT COUNT(*) FROM [@PS_GA150H] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "DELETE FROM [@PS_GA150H] WHERE DocEntry = '" + DocEntry + "'";
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
		/// 기본정보 수정
		/// </summary>
		/// <returns></returns>
		private bool PS_GA150_UpdateData()
		{
			bool ReturnValue = false;
			int DocEntry;
			string BPLID;    //사업장
			string CntcCode; //사번
			string CntcName; //성명
			string TeamCd;   //팀
			string TeamNm;   //팀명
			string RspCd;    //담당
			string RspNm;    //담당명
			string PrtrCd;   //보호구코드
			string PrtrNm;   //보호구명
			string PrtrSpec; //보호구규격
			string PrvdDt;   //지급일자
			Double Pred;     //주기
			Double Qty;      //수량
			string RtrnDt;   //반납일자
			string RtrnRsn;  //반납사유
			string Note;     //비고
			string RegDt;    //등록일
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				TeamCd = oForm.Items.Item("TeamCd").Specific.Value.ToString().Trim();
				TeamNm = oForm.Items.Item("TeamNm").Specific.Value.ToString().Trim();
				RspCd = oForm.Items.Item("RspCd").Specific.Value.ToString().Trim();
				RspNm = oForm.Items.Item("RspNm").Specific.Value.ToString().Trim();
				PrtrCd = oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim();
				PrtrNm = oForm.Items.Item("PrtrNm").Specific.Value.ToString().Trim();
				PrtrSpec = oForm.Items.Item("PrtrSpec").Specific.Value.ToString().Trim();
				PrvdDt = oForm.Items.Item("PrvdDt").Specific.Value.ToString().Trim();
				Pred = Convert.ToDouble(oForm.Items.Item("Pred").Specific.Value.ToString().Trim());
				Qty = Convert.ToDouble(oForm.Items.Item("Qty").Specific.Value.ToString().Trim());
				RtrnDt = oForm.Items.Item("RtrnDt").Specific.Value.ToString().Trim();
				RtrnRsn = oForm.Items.Item("RtrnRsn").Specific.Value.ToString().Trim();
				Note = oForm.Items.Item("Note").Specific.Value.ToString().Trim();
				RegDt = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(Convert.ToString(DocEntry)))
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요.";
					throw new Exception();
				}

				sQry = " UPDATE   [@PS_GA150H]";
				sQry += " SET      U_BPLId = '" + BPLID + "',";
				sQry += "          U_CntcCode = '" + CntcCode + "',";
				sQry += "          U_CntcName = '" + CntcName + "',";
				sQry += "          U_TeamCd = '" + TeamCd + "',";
				sQry += "          U_TeamNm = '" + TeamNm + "',";
				sQry += "          U_RspCd = '" + RspCd + "',";
				sQry += "          U_RspNm = '" + RspNm + "',";
				sQry += "          U_PrtrCd = '" + PrtrCd + "',";
				sQry += "          U_PrtrNm = '" + PrtrNm + "',";
				sQry += "          U_PrtrSpec = '" + PrtrSpec + "',";
				sQry += "          U_PrvdDt = '" + PrvdDt + "',";
				sQry += "          U_Pred = '" + Pred + "',";
				sQry += "          U_Qty = '" + Qty + "',";
				sQry += "          U_RtrnDt = '" + RtrnDt + "',";
				sQry += "          U_RtrnRsn = '" + RtrnRsn + "',";
				sQry += "          U_Note = '" + Note + "',";
				sQry += "          U_RegDt = '" + RegDt + "'";
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
		private bool PS_GA150_AddData()
		{
			bool ReturnValue = false;
			int DocEntry;
			string BPLID;    //사업장
			string CntcCode; //사번
			string CntcName; //성명
			string TeamCd;   //팀
			string TeamNm;   //팀명
			string RspCd;    //담당
			string RspNm;    //담당명
			string PrtrCd;   //보호구코드
			string PrtrNm;   //보호구명
			string PrtrSpec; //보호구규격
			string PrvdDt;   //지급일자
			string LPrvdDt;  //전지급일
			Double Pred;        //주기
			Double Qty;         //수량
			string RtrnDt;   //반납일자
			string RtrnRsn;  //반납사유
			string Note;     //비고
			string RegDt;    //등록일
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				TeamCd = oForm.Items.Item("TeamCd").Specific.Value.ToString().Trim();
				TeamNm = oForm.Items.Item("TeamNm").Specific.Value.ToString().Trim();
				RspCd = oForm.Items.Item("RspCd").Specific.Value.ToString().Trim();
				RspNm = oForm.Items.Item("RspNm").Specific.Value.ToString().Trim();
				PrtrCd = oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim();
				PrtrNm = oForm.Items.Item("PrtrNm").Specific.Value.ToString().Trim();
				PrtrSpec = oForm.Items.Item("PrtrSpec").Specific.Value.ToString().Trim();
				PrvdDt = oForm.Items.Item("PrvdDt").Specific.Value.ToString().Trim();
				LPrvdDt = oForm.Items.Item("LPrvdDt").Specific.Value.ToString().Trim();
				Pred = Convert.ToDouble(oForm.Items.Item("Pred").Specific.Value.ToString().Trim());
				Qty = Convert.ToDouble(oForm.Items.Item("Qty").Specific.Value.ToString().Trim());
				RtrnDt = oForm.Items.Item("RtrnDt").Specific.Value.ToString().Trim();
				RtrnRsn = oForm.Items.Item("RtrnRsn").Specific.Value.ToString().Trim();
				Note = oForm.Items.Item("Note").Specific.Value.ToString().Trim();
				RegDt = oForm.Items.Item("RegDt").Specific.Value.ToString().Trim();

				//DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_GA150H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = 1;
				}
				else
				{
					DocEntry = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}

				sQry = "  INSERT INTO [@PS_GA150H]";
				sQry += " (";
				sQry += "     DocEntry,";
				sQry += "     DocNum,";
				sQry += "     U_BPLId,";
				sQry += "     U_CntcCode,";
				sQry += "     U_CntcName,";
				sQry += "     U_TeamCd,";
				sQry += "     U_TeamNm,";
				sQry += "     U_RspCd,";
				sQry += "     U_RspNm,";
				sQry += "     U_PrtrCd,";
				sQry += "     U_PrtrNm,";
				sQry += "     U_PrtrSpec,";
				sQry += "     U_PrvdDt,";
				sQry += "     U_LPrvdDt,";
				sQry += "     U_Pred,";
				sQry += "     U_Qty,";
				sQry += "     U_RtrnDt,";
				sQry += "     U_RtrnRsn,";
				sQry += "     U_Note,";
				sQry += "     U_RegDt";
				sQry += " )";
				sQry += " VALUES";
				sQry += " (";
				sQry += DocEntry + ",";
				sQry += DocEntry + ",";
				sQry += "'" + BPLID + "',";
				sQry += "'" + CntcCode + "',";
				sQry += "'" + CntcName + "',";
				sQry += "'" + TeamCd + "',";
				sQry += "'" + TeamNm + "',";
				sQry += "'" + RspCd + "',";
				sQry += "'" + RspNm + "',";
				sQry += "'" + PrtrCd + "',";
				sQry += "'" + PrtrNm + "',";
				sQry += "'" + PrtrSpec + "',";
				sQry += "'" + PrvdDt + "',";
				sQry += "'" + LPrvdDt + "',";
				sQry += "'" + Pred + "',";
				sQry += "'" + Qty + "',";
				sQry += "'" + RtrnDt + "',";
				sQry += "'" + RtrnRsn + "',";
				sQry += "'" + Note + "',";
				sQry += "'" + RegDt + "'";
				sQry += ")";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// 필수입력사항 체크
		/// </summary>
		/// <returns></returns>
		private bool PS_GA150_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "사번은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim()))
				{
					errMessage = "보호구코드는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("PrvdDt").Specific.Value.ToString().Trim()))
				{
					errMessage = "지급일자는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				else if (Convert.ToDouble(oForm.Items.Item("Qty").Specific.Value.ToString().Trim()) == 0)
				{
					errMessage = "수량은 필수사항입니다. 확인하세요.";
					throw new Exception();
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
			return ReturnValue;
		}

		/// <summary>
		///  부자재 청구서 출력 시 필수입력사항 체크
		/// </summary>
		/// <returns></returns>
		private bool PS_GA150_RegDtCheck()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("SRegDt").Specific.Value.ToString().Trim()))
				{
					errMessage = "부자재청구서 출력시에는 등록일을 필수로 입력해야합니다. 확인하세요.";
					throw new Exception();
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
			return ReturnValue;
		}

		/// <summary>
		/// PS_GA150_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_GA150_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLID; //사업장
			string RegDt; //등록일
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("SBPLId").Specific.Selected.Value.ToString().Trim();
				RegDt = oForm.Items.Item("SRegDt").Specific.Value.ToString().Trim();

				WinTitle = "[PS_GA150] 레포트";
				ReportName = "PS_GA150_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@RegDt", RegDt));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
							if (PS_GA150_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_GA150_DataCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_GA150_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_GA150_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA150_LoadCaption();
							PS_GA150_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_GA150_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_GA150_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_GA150_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA150_LoadCaption();
							PS_GA150_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						PS_GA150_FormReset();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_GA150_LoadCaption();
						PS_GA150_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_GA150_DeleteData();
							PS_GA150_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA150_LoadCaption();
							PS_GA150_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnPrt")
					{
						if (PS_GA150_RegDtCheck() == false)
						{
							BubbleEvent = false;
							return;
						}
						System.Threading.Thread thread = new System.Threading.Thread(PS_GA150_Print_Report01);
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
							oDS_PS_GA150H.SetValue("DocEntry", 0, oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_BPLId", 0, oMat.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_CntcCode", 0, oMat.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_CntcName", 0, oMat.Columns.Item("CntcName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_TeamCd", 0, oMat.Columns.Item("TeamCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_TeamNm", 0, oMat.Columns.Item("TeamNm").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_RspCd", 0, oMat.Columns.Item("RspCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_RspNm", 0, oMat.Columns.Item("RspNm").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_PrtrCd", 0, oMat.Columns.Item("PrtrCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_PrtrNm", 0, oMat.Columns.Item("PrtrNm").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_PrtrSpec", 0, oMat.Columns.Item("PrtrSpec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_PrvdDt", 0, oMat.Columns.Item("PrvdDt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_LPrvdDt", 0, oMat.Columns.Item("LPrvdDt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_Pred", 0, oMat.Columns.Item("Pred").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_Qty", 0, oMat.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_RtrnDt", 0, oMat.Columns.Item("RtrnDt").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));
							oDS_PS_GA150H.SetValue("U_RtrnRsn", 0, oMat.Columns.Item("RtrnRsn").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_Note", 0, oMat.Columns.Item("Note").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							oDS_PS_GA150H.SetValue("U_RegDt", 0, oMat.Columns.Item("RegDt").Cells.Item(pVal.Row).Specific.Value.Replace(".", ""));

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_GA150_LoadCaption();
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							oForm.Items.Item("CntcName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("TeamCd").Specific.Value = dataHelpClass.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("RspCd").Specific.Value = dataHelpClass.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("TeamNm").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("TeamCd").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("RspNm").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("RspCd").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("LPrvdDt").Specific.Value = dataHelpClass.Get_ReData("Convert(VARCHAR(10), MAX(U_PrvdDt), 112) ", "U_CntcCode", "[@PS_GA150H]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "AND U_PrtrNm = '" + oForm.Items.Item("PrtrNm").Specific.Value.ToString().Trim() + "'");
						}
						else if (pVal.ItemUID == "PrtrCd")
						{
							oForm.Items.Item("PrtrNm").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("PrtrSpec").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item("PrtrCd").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("LPrvdDt").Specific.Value = dataHelpClass.Get_ReData("Convert(VARCHAR(10), MAX(U_PrvdDt), 112) ", "U_CntcCode", "[@PS_GA150H]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "AND U_PrtrNm = '" + oForm.Items.Item("PrtrNm").Specific.Value.ToString().Trim() + "'");
						}
						else if (pVal.ItemUID == "SCntcCode")
						{
							oForm.Items.Item("SCntcName").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SCntcCode").Specific.Value.ToString().Trim() + "'", "");
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA150H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA150L);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							PS_GA150_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_GA150_LoadCaption();
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
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
