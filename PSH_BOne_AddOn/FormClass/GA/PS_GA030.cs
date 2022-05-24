using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 예산 등록
	/// </summary>
	internal class PS_GA030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_GA030L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_GA030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_GA030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_GA030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_GA030_CreateItems();
				PS_GA030_ComboBox_Setting();
				PS_GA030_LoadCaption();

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
		/// PS_GA030_CreateItems
		/// </summary>
		private void PS_GA030_CreateItems()
		{
			try
			{
				oDS_PS_GA030L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");

				//계정
				oForm.DataSources.UserDataSources.Add("AcctCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("AcctCode01").Specific.DataBind.SetBound(true, "", "AcctCode01");

				//계정명
				oForm.DataSources.UserDataSources.Add("AcctName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("AcctName01").Specific.DataBind.SetBound(true, "", "AcctName01");

				//계정과목
				oForm.DataSources.UserDataSources.Add("AcctCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("AcctCode02").Specific.DataBind.SetBound(true, "", "AcctCode02");

				//계정과목명
				oForm.DataSources.UserDataSources.Add("AcctName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("AcctName02").Specific.DataBind.SetBound(true, "", "AcctName02");

				//세부계정과목
				oForm.DataSources.UserDataSources.Add("AcctCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("AcctCode03").Specific.DataBind.SetBound(true, "", "AcctCode03");

				//세부계정과목명
				oForm.DataSources.UserDataSources.Add("AcctName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("AcctName03").Specific.DataBind.SetBound(true, "", "AcctName03");

				//비율
				oForm.DataSources.UserDataSources.Add("Rate", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("Rate").Specific.DataBind.SetBound(true, "", "Rate");

				oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");

				//년도, 사업장 코드는 Hidden 처리
				oMat.Columns.Item("StdYear").Visible = false;
				oMat.Columns.Item("BPLID").Visible = false;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_GA030_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_GA030_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
		/// 메트릭스 Row추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_GA030_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_GA030L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_GA030L.Offset = oRow;
				oDS_PS_GA030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 데이터 조회
		/// </summary>
		/// <param name="pSearchType"></param>
		private void PS_GA030_MTX01(string pSearchType)
		{
			int loopCount;
			string BPLID;      //사업장
			string StdYear;    //기준년도
			string AcctCode01; //계정코드
			string AcctCode02; //계정과목코드
			string AcctCode03; //세부계정과목코드
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				AcctCode01 = oForm.Items.Item("AcctCode01").Specific.Value.ToString().Trim();
				AcctCode02 = oForm.Items.Item("AcctCode02").Specific.Value.ToString().Trim();
				AcctCode03 = oForm.Items.Item("AcctCode03").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				if (pSearchType == "1") //실적조회
				{
					sQry = " EXEC [PS_GA030_01] '";
					sQry += BPLID + "','";
					sQry += StdYear + "','";
					sQry += AcctCode01 + "','";
					sQry += AcctCode02 + "','";
					sQry += AcctCode03 + "'";
				}
				else //계획조회
				{
					sQry = " EXEC [PS_GA030_02] '";
					sQry += BPLID + "','";
					sQry += StdYear + "','";
					sQry += AcctCode01 + "','";
					sQry += AcctCode02 + "','";
					sQry += AcctCode03 + "'";
				}
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_GA030L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_GA030_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount + 1 > oDS_PS_GA030L.Size)
					{
						oDS_PS_GA030L.InsertRecord(loopCount);
					}

					oMat.AddRow();
					oDS_PS_GA030L.Offset = loopCount;
					oDS_PS_GA030L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
					oDS_PS_GA030L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("StdYear").Value.ToString().Trim());     //년도
					oDS_PS_GA030L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("BPLID").Value.ToString().Trim());       //사업장코드
					oDS_PS_GA030L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("AcctCode01").Value.ToString().Trim());  //계정코드
					oDS_PS_GA030L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("AcctName01").Value.ToString().Trim());  //계정명
					oDS_PS_GA030L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("AcctCode02").Value.ToString().Trim());  //계정과목코드
					oDS_PS_GA030L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("AcctName02").Value.ToString().Trim());  //계정과목명
					oDS_PS_GA030L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("AcctCode03").Value.ToString().Trim());  //세부계정과목코드
					oDS_PS_GA030L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("AcctName03").Value.ToString().Trim());  //세부계정과목명
					oDS_PS_GA030L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("Month01").Value.ToString().Trim());     //1월
					oDS_PS_GA030L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("Month02").Value.ToString().Trim());     //2월
					oDS_PS_GA030L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Month03").Value.ToString().Trim());     //3월
					oDS_PS_GA030L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("Month04").Value.ToString().Trim());     //4월
					oDS_PS_GA030L.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("Month05").Value.ToString().Trim());     //5월
					oDS_PS_GA030L.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("Month06").Value.ToString().Trim());     //6월
					oDS_PS_GA030L.SetValue("U_ColSum07", loopCount, oRecordSet.Fields.Item("Month07").Value.ToString().Trim());     //7월
					oDS_PS_GA030L.SetValue("U_ColSum08", loopCount, oRecordSet.Fields.Item("Month08").Value.ToString().Trim());     //8월
					oDS_PS_GA030L.SetValue("U_ColSum09", loopCount, oRecordSet.Fields.Item("Month09").Value.ToString().Trim());     //9월
					oDS_PS_GA030L.SetValue("U_ColSum10", loopCount, oRecordSet.Fields.Item("Month10").Value.ToString().Trim());     //10월
					oDS_PS_GA030L.SetValue("U_ColSum11", loopCount, oRecordSet.Fields.Item("Month11").Value.ToString().Trim());     //11월
					oDS_PS_GA030L.SetValue("U_ColSum12", loopCount, oRecordSet.Fields.Item("Month12").Value.ToString().Trim());     //12월
					oDS_PS_GA030L.SetValue("U_ColSum13", loopCount, oRecordSet.Fields.Item("Total").Value.ToString().Trim());       //계
					oDS_PS_GA030L.SetValue("U_ColPrc01", loopCount, oRecordSet.Fields.Item("Rate").Value.ToString().Trim());        //비율
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
		///  실적*비율=예산 (전체 계산)
		/// </summary>
		private void PS_GA030_Calculate_Header()
		{
			int loopCount;
			double rate_Renamed;

			try
			{
				oForm.Freeze(true);
				rate_Renamed =	Convert.ToDouble(oForm.Items.Item("Rate").Specific.Value.ToString().Trim()) * 0.01;
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.RowCount - 1; loopCount++)
				{
					oDS_PS_GA030L.SetValue("U_ColSum01", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount).ToString().Trim())) * rate_Renamed)); //1월
					oDS_PS_GA030L.SetValue("U_ColSum02", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount).ToString().Trim())) * rate_Renamed)); //2월
					oDS_PS_GA030L.SetValue("U_ColSum03", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount).ToString().Trim())) * rate_Renamed)); //3월
					oDS_PS_GA030L.SetValue("U_ColSum04", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount).ToString().Trim())) * rate_Renamed)); //4월
					oDS_PS_GA030L.SetValue("U_ColSum05", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount).ToString().Trim())) * rate_Renamed)); //5월
					oDS_PS_GA030L.SetValue("U_ColSum06", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount).ToString().Trim())) * rate_Renamed)); //6월
					oDS_PS_GA030L.SetValue("U_ColSum07", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount).ToString().Trim())) * rate_Renamed)); //7월
					oDS_PS_GA030L.SetValue("U_ColSum08", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount).ToString().Trim())) * rate_Renamed)); //8월
					oDS_PS_GA030L.SetValue("U_ColSum09", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount).ToString().Trim())) * rate_Renamed)); //9월
					oDS_PS_GA030L.SetValue("U_ColSum10", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount).ToString().Trim())) * rate_Renamed)); //10월
					oDS_PS_GA030L.SetValue("U_ColSum11", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount).ToString().Trim())) * rate_Renamed)); //11월
					oDS_PS_GA030L.SetValue("U_ColSum12", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount).ToString().Trim())) * rate_Renamed)); //12월
					oDS_PS_GA030L.SetValue("U_ColSum13", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount).ToString().Trim())) * rate_Renamed)); //계
					oDS_PS_GA030L.SetValue("U_ColPrc01", loopCount, Convert.ToString(rate_Renamed * 100));  //비율(%)
				}

				oMat.LoadFromDataSource();
				PSH_Globals.SBO_Application.StatusBar.SetText("계산 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// 실적*비율=예산 (라인별 계산)
		/// </summary>
		private void PS_GA030_Calculate_Line()
		{
			int loopCount;
			double rate_Renamed;

			try
			{
				oForm.Freeze(true);
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.RowCount - 1; loopCount++)
				{
					rate_Renamed = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount).ToString().Trim()) * 0.01;
					oDS_PS_GA030L.SetValue("U_ColSum01", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount).ToString().Trim())) * rate_Renamed)); //1월
					oDS_PS_GA030L.SetValue("U_ColSum02", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount).ToString().Trim())) * rate_Renamed)); //2월
					oDS_PS_GA030L.SetValue("U_ColSum03", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount).ToString().Trim())) * rate_Renamed)); //3월
					oDS_PS_GA030L.SetValue("U_ColSum04", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount).ToString().Trim())) * rate_Renamed)); //4월
					oDS_PS_GA030L.SetValue("U_ColSum05", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount).ToString().Trim())) * rate_Renamed)); //5월
					oDS_PS_GA030L.SetValue("U_ColSum06", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount).ToString().Trim())) * rate_Renamed)); //6월
					oDS_PS_GA030L.SetValue("U_ColSum07", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount).ToString().Trim())) * rate_Renamed)); //7월
					oDS_PS_GA030L.SetValue("U_ColSum08", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount).ToString().Trim())) * rate_Renamed)); //8월
					oDS_PS_GA030L.SetValue("U_ColSum09", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount).ToString().Trim())) * rate_Renamed)); //9월
					oDS_PS_GA030L.SetValue("U_ColSum10", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount).ToString().Trim())) * rate_Renamed)); //10월
					oDS_PS_GA030L.SetValue("U_ColSum11", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount).ToString().Trim())) * rate_Renamed)); //11월
					oDS_PS_GA030L.SetValue("U_ColSum12", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount).ToString().Trim())) * rate_Renamed)); //12월
					oDS_PS_GA030L.SetValue("U_ColSum13", loopCount, Convert.ToString((Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount).ToString().Trim()) + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount).ToString().Trim())) * rate_Renamed)); //계
					oDS_PS_GA030L.SetValue("U_ColPrc01", loopCount, oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount).ToString().Trim());               //비율
				}

				oMat.LoadFromDataSource();
				PSH_Globals.SBO_Application.StatusBar.SetText("개별 계산 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// 데이터 INSERT
		/// </summary>
		private void PS_GA030_AddData()
		{
			int loopCount;
			string StdYear;	   //년도
			string BPLID;	   //사업장
			string AcctCode01; //계정코드
			string AcctName01; //계정명
			string AcctCode02; //계정과목코드
			string AcctName02; //계정과목명
			string AcctCode03; //세부계정과목코드
			string AcctName03; //세부계정과목명
			double Month01;	   //01월
			double Month02;	   //02월
			double Month03;	   //03월
			double Month04;	   //04월
			double Month05;	   //05월
			double Month06;	   //06월
			double Month07;	   //07월
			double Month08;	   //08월
			double Month09;	   //09월
			double Month10;	   //10월
			double Month11;	   //11월
			double Month12;	   //12월
			double Total;	   //계
			double rate_Renamed; //비율
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.RowCount - 1; loopCount++)
				{

					StdYear = oDS_PS_GA030L.GetValue("U_ColReg01", loopCount).ToString().Trim();
					BPLID = oDS_PS_GA030L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					AcctCode01 = oDS_PS_GA030L.GetValue("U_ColReg03", loopCount).ToString().Trim();
					AcctName01 = oDS_PS_GA030L.GetValue("U_ColReg04", loopCount).ToString().Trim();
					AcctCode02 = oDS_PS_GA030L.GetValue("U_ColReg05", loopCount).ToString().Trim();
					AcctName02 = oDS_PS_GA030L.GetValue("U_ColReg06", loopCount).ToString().Trim();
					AcctCode03 = oDS_PS_GA030L.GetValue("U_ColReg07", loopCount).ToString().Trim();
					AcctName03 = oDS_PS_GA030L.GetValue("U_ColReg08", loopCount).ToString().Trim();
					Month01 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount).ToString().Trim());
					Month02 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount).ToString().Trim());
					Month03 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount).ToString().Trim());
					Month04 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount).ToString().Trim());
					Month05 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount).ToString().Trim());
					Month06 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount).ToString().Trim());
					Month07 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount).ToString().Trim());
					Month08 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount).ToString().Trim());
					Month09 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount).ToString().Trim());
					Month10 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount).ToString().Trim());
					Month11 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount).ToString().Trim());
					Month12 = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount).ToString().Trim());
					Total = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount).ToString().Trim());
					rate_Renamed = Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount).ToString().Trim());

					sQry = "     EXEC [PS_GA030_03] ";
					sQry += "'" + StdYear + "',";
					sQry += "'" + BPLID + "',";
					sQry += "'" + AcctCode01 + "',";
					sQry += "'" + AcctName01 + "',";
					sQry += "'" + AcctCode02 + "',";
					sQry += "'" + AcctName02 + "',";
					sQry += "'" + AcctCode03 + "',";
					sQry += "'" + AcctName03 + "',";
					sQry += "'" + Month01 + "',";
					sQry += "'" + Month02 + "',";
					sQry += "'" + Month03 + "',";
					sQry += "'" + Month04 + "',";
					sQry += "'" + Month05 + "',";
					sQry += "'" + Month06 + "',";
					sQry += "'" + Month07 + "',";
					sQry += "'" + Month08 + "',";
					sQry += "'" + Month09 + "',";
					sQry += "'" + Month10 + "',";
					sQry += "'" + Month11 + "',";
					sQry += "'" + Month12 + "',";
					sQry += "'" + Total + "',";
					sQry += "'" + rate_Renamed + "'";
					oRecordSet.DoQuery(sQry);
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_GA030_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_GA030_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string Total;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						oMat.FlushToDataSource();

						if (oCol == "Month01" || oCol == "Month02" || oCol == "Month03" || oCol == "Month04" || oCol == "Month05" 
							         || oCol == "Month06" || oCol == "Month07" || oCol == "Month08" || oCol == "Month09" || oCol == "Month10"
									 || oCol == "Month11" || oCol == "Month12")
						{
							Total = Convert.ToString(Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum01", oRow - 1).ToString().Trim())
								                     + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum02", oRow - 1).ToString().Trim())
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum03", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum04", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum05", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum06", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum07", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum08", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum09", oRow - 1).ToString().Trim())
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum10", oRow - 1).ToString().Trim()) 
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum11", oRow - 1).ToString().Trim())
													 + Convert.ToDouble(oDS_PS_GA030L.GetValue("U_ColSum12", oRow - 1).ToString().Trim()));
							oDS_PS_GA030L.SetValue("U_ColSum13", oRow - 1, Total);
							oMat.LoadFromDataSource();
							oMat.Columns.Item(oCol).Cells.Item(oRow).Click();
						}
						oMat.AutoResizeColumns();
						break;
					case "AcctCode01": //계정
						oForm.Items.Item("AcctName01").Specific.Value = dataHelpClass.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" + oForm.Items.Item("AcctCode01").Specific.Value.ToString().Trim() + "'", ""); 
						break;
					case "AcctCode02": //계정과목
						oForm.Items.Item("AcctName02").Specific.Value = dataHelpClass.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" + oForm.Items.Item("AcctCode02").Specific.Value.ToString().Trim() + "'", "");
						break;
					case "AcctCode03": //세부계정과목
						oForm.Items.Item("AcctName03").Specific.Value = dataHelpClass.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" + oForm.Items.Item("AcctCode03").Specific.Value.ToString().Trim() + "'", "");
						break;
				}
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
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
							PS_GA030_AddData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA030_LoadCaption();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							PS_GA030_AddData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA030_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnSearch1")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_GA030_LoadCaption();
						PS_GA030_MTX01("1");
					}
					else if (pVal.ItemUID == "BtnSearch2")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_GA030_LoadCaption();
						PS_GA030_MTX01("2");
					}
					else if (pVal.ItemUID == "BtnCal1")
					{
						PS_GA030_Calculate_Header();
					}
					else if (pVal.ItemUID == "BtnCal2")
					{
						PS_GA030_Calculate_Line();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode03", "");
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
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
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "Month01" || pVal.ColUID == "Month02" || pVal.ColUID == "Month03" || pVal.ColUID == "Month04" || pVal.ColUID == "Month05" 
								|| pVal.ColUID == "Month06" || pVal.ColUID == "Month07" || pVal.ColUID == "Month08" || pVal.ColUID == "Month09" || pVal.ColUID == "Month10" 
								|| pVal.ColUID == "Month11" || pVal.ColUID == "Month12")
							{
								PS_GA030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
						}
						else
						{
							PS_GA030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA030L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_GA030L.RemoveRecord(oDS_PS_GA030L.Size - 1);
						oMat.LoadFromDataSource();
					}
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
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_GA030_LoadCaption();
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "7169": //엑셀 내보내기
							PS_GA030_Add_MatrixRow(oMat.VisualRowCount, false);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "7169": //엑셀 내보내기
							oDS_PS_GA030L.RemoveRecord(oDS_PS_GA030L.Size - 1);
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
