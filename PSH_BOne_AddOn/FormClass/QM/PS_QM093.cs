using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 실패비용 직접 등록 및 조회
	/// </summary>
	internal class PS_QM093 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM093H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM093L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM093.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM093_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM093");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM093_CreateItems();
				PS_QM093_ComboBox_Setting();
				PS_QM093_EnableMenus();
				PS_QM093_FormResize();
				PS_QM093_LoadCaption();
				PS_QM093_FormReset();
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
		/// PS_QM093_CreateItems
		/// </summary>
		private void PS_QM093_CreateItems()
		{
			try
			{
				oDS_PS_QM093H = oForm.DataSources.DBDataSources.Item("@PS_QM093H");
				oDS_PS_QM093L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//관리번호
				oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

				//사업장
				oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

				//작번
				oForm.DataSources.UserDataSources.Add("SItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SItemCode").Specific.DataBind.SetBound(true, "", "SItemCode");

				//품명
				oForm.DataSources.UserDataSources.Add("SItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("SItemName").Specific.DataBind.SetBound(true, "", "SItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("SItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("SItemSpec").Specific.DataBind.SetBound(true, "", "SItemSpec");

				//기간(FR)
				oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

				//기간(TO)
				oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

				oMat.Columns.Item("Check").Visible = false; //선택 체크박스 Visible = False

				//SET
				oForm.Items.Item("SFrDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("SToDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("ItemCode").Click(); //작번 포커스
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_QM093_ComboBox_Setting()
		{
			string User_BPLId;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				User_BPLId = dataHelpClass.User_BPLID();

				//기본정보 사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("CLTCOD").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);

				//조회정보 사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("SCLTCOD").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);

				//Line
				sQry = " SELECT      BPLId, ";
				sQry += "             BPLName ";
				sQry += " FROM        [OBPL] ";
				sQry += " ORDER BY    BPLId";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("CLTCOD"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM093_EnableMenus
		/// </summary>
		private void PS_QM093_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM093_FormResize
		/// </summary>
		private void PS_QM093_FormResize()
		{
			try
			{
				oForm.Items.Item("GrpBox01").Height = 310;
				oForm.Items.Item("GrpBox01").Width = 590;

				oForm.Items.Item("GrpBox02").Height = 105;
				oForm.Items.Item("GrpBox02").Width = 390;

				oForm.Items.Item("GrpBox03").Height = 165;
				oForm.Items.Item("GrpBox03").Width = 260;

				oForm.Items.Item("GrpBox04").Height = 165;
				oForm.Items.Item("GrpBox04").Width = 225;

				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM093_LoadCaption
		/// </summary>
		private void PS_QM093_LoadCaption()
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
		/// PS_QM093_FormReset
		/// </summary>
		private void PS_QM093_FormReset()
		{
			string User_BPLId;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//관리번호
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_QM093H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oDS_PS_QM093H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_QM093H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
				}

				User_BPLId = dataHelpClass.User_BPLID();

				//기준정보
				oDS_PS_QM093H.SetValue("U_CLTCOD", 0, User_BPLId);  //사업장
				oDS_PS_QM093H.SetValue("U_ItemCode", 0, "");	    //작번
				oDS_PS_QM093H.SetValue("U_ItemName", 0, "");	    //품명
				oDS_PS_QM093H.SetValue("U_ItemSpec", 0, "");	    //규격
				oDS_PS_QM093H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd")); //날짜
				oDS_PS_QM093H.SetValue("U_TotalQty", 0, "0");	    //전체수량
				oDS_PS_QM093H.SetValue("U_BadQty", 0, "0");			//불량수량
				oDS_PS_QM093H.SetValue("U_BadNote", 0, "");			//불량내용
				oDS_PS_QM093H.SetValue("U_CpCode", 0, "");			//원인공정
				oDS_PS_QM093H.SetValue("U_CpName", 0, "");			//작업자
				oDS_PS_QM093H.SetValue("U_WorkCode", 0, "");		//작업자명
				oDS_PS_QM093H.SetValue("U_WorkName", 0, "");		//최종판정
				oDS_PS_QM093H.SetValue("U_Cost01", 0, "");			//자재
				oDS_PS_QM093H.SetValue("U_Cost02", 0, "");			//가공
				oDS_PS_QM093H.SetValue("U_Cost03", 0, "");			//설계
				oDS_PS_QM093H.SetValue("U_Cost04", 0, "");			//외주
				oDS_PS_QM093H.SetValue("U_Cost05", 0, "");			//분해조립
				oDS_PS_QM093H.SetValue("U_Cost06", 0, "");			//A/S출장
				oDS_PS_QM093H.SetValue("U_Cost07", 0, "");			//운송
				oDS_PS_QM093H.SetValue("U_Cost08", 0, "");			//지체상금
				oDS_PS_QM093H.SetValue("U_CostTot", 0, "");			//계
				oDS_PS_QM093H.SetValue("U_Comments", 0, "");		//비고

				oForm.Items.Item("ItemCode").Click();
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
		/// 합계 금액 계산
		/// </summary>
		private void PS_QM093_CalculateTotalExp()
		{
			decimal Cost01;	 //자재
			decimal Cost02;	 //가공
			decimal Cost03;	 //설계
			decimal Cost04;	 //외주
			decimal Cost05;	 //분해조립
			decimal Cost06;	 //A/S출장
			decimal Cost07;	 //운송
			decimal Cost08;	 //지체상금
			decimal CostTot; //합계

			try
			{
				Cost01 = Convert.ToDecimal(oForm.Items.Item("Cost01").Specific.Value.ToString().Trim());
				Cost02 = Convert.ToDecimal(oForm.Items.Item("Cost02").Specific.Value.ToString().Trim());
				Cost03 = Convert.ToDecimal(oForm.Items.Item("Cost03").Specific.Value.ToString().Trim());
				Cost04 = Convert.ToDecimal(oForm.Items.Item("Cost04").Specific.Value.ToString().Trim());
				Cost05 = Convert.ToDecimal(oForm.Items.Item("Cost05").Specific.Value.ToString().Trim());
				Cost06 = Convert.ToDecimal(oForm.Items.Item("Cost06").Specific.Value.ToString().Trim());
				Cost07 = Convert.ToDecimal(oForm.Items.Item("Cost07").Specific.Value.ToString().Trim());
				Cost08 = Convert.ToDecimal(oForm.Items.Item("Cost08").Specific.Value.ToString().Trim());
				CostTot = Cost01 + Cost02 + Cost03 + Cost04 + Cost05 + Cost06 + Cost07 + Cost08;

				oDS_PS_QM093H.SetValue("U_CostTot", 0, Convert.ToString(CostTot));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 메트릭스 Row추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM093_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_QM093L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM093L.Offset = oRow;
				oDS_PS_QM093L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		private void PS_QM093_MTX01()
		{
			int i;
			string SDocEntry; //관리번호
			string SCLTCOD;	  //사업장
			string SItemCode; //작번
			string SFrDate;	  //기간(FR)
			string SToDate;   //기간(TO)
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				SDocEntry = oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim();
				SCLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
				SItemCode = oForm.Items.Item("SItemCode").Specific.Value.ToString().Trim();
				SFrDate = oForm.Items.Item("SFrDate").Specific.Value.ToString().Trim();
				SToDate = oForm.Items.Item("SToDate").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "     EXEC [PS_QM093_01] ";
				sQry += "'" + SDocEntry + "',";
				sQry += "'" + SCLTCOD + "',";
				sQry += "'" + SItemCode + "',";
				sQry += "'" + SFrDate + "',";
				sQry += "'" + SToDate + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM093L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_QM093_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM093L.Size)
					{
						oDS_PS_QM093L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_QM093L.Offset = i;

					oDS_PS_QM093L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM093L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim()); //관리번호
					oDS_PS_QM093L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CLTCOD").Value.ToString().Trim());	 //사업장
					oDS_PS_QM093L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //작번
					oDS_PS_QM093L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
					oDS_PS_QM093L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim()); //규격
					oDS_PS_QM093L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd"));	//날짜
					oDS_PS_QM093L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("TotalQty").Value.ToString().Trim()); //전체수량
					oDS_PS_QM093L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("BadQty").Value.ToString().Trim());	 //불량수량
					oDS_PS_QM093L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("BadNote").Value.ToString().Trim());	 //불량내용
					oDS_PS_QM093L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());	 //원인공정
					oDS_PS_QM093L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());	 //원인공정명
					oDS_PS_QM093L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("WorkCode").Value.ToString().Trim()); //작업자
					oDS_PS_QM093L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("WorkName").Value.ToString().Trim()); //작업자명
					oDS_PS_QM093L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("LastNote").Value.ToString().Trim()); //최종판정
					oDS_PS_QM093L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Cost01").Value.ToString().Trim());	 //자재
					oDS_PS_QM093L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Cost02").Value.ToString().Trim());	 //가공
					oDS_PS_QM093L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("Cost03").Value.ToString().Trim());	 //설계
					oDS_PS_QM093L.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("Cost04").Value.ToString().Trim());	 //외주
					oDS_PS_QM093L.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("Cost05").Value.ToString().Trim());	 //분해조립
					oDS_PS_QM093L.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("Cost06").Value.ToString().Trim());	 //A/S출장
					oDS_PS_QM093L.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("Cost07").Value.ToString().Trim());	 //운송
					oDS_PS_QM093L.SetValue("U_ColSum08", i, oRecordSet.Fields.Item("Cost08").Value.ToString().Trim());	 //지체상금
					oDS_PS_QM093L.SetValue("U_ColSum09", i, oRecordSet.Fields.Item("CostTot").Value.ToString().Trim());	 //계
					oDS_PS_QM093L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("Comments").Value.ToString().Trim()); //비고
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
		private void PS_QM093_DeleteData()
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

					sQry = "SELECT COUNT(*) FROM [@PS_QM093H] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "EXEC PS_QM093_04 '" + DocEntry + "'";
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
		/// 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_QM093_AddData()
		{
			bool ReturnValue = false;
			int DocEntry;	 //관리번호
			string CLTCOD;	 //사업장
			string ItemCode; //작번
			string ItemName; //품명
			string ItemSpec; //규격
			string DocDate;	 //날짜
			int TotalQty;	 //전체수량
			int BadQty;		 //불량수량
			string BadNote;	 //불량내용
			string CpCode;	 //원인공정
			string CpName;	 //원인공정명
			string WorkCode; //작업자
			string WorkName; //작업자명
			string LastNote; //최종판정
			decimal Cost01;	 //자재
			decimal Cost02;	 //가공
			decimal Cost03;	 //설계
			decimal Cost04;	 //외주
			decimal Cost05;	 //분해조립
			decimal Cost06;	 //A/S출장
			decimal Cost07;	 //운송
			decimal Cost08;	 //지체상금
			decimal CostTot; //계
			string Comments; //비고
			string UserSign; //UserSign
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				ItemSpec = oForm.Items.Item("ItemSpec").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				TotalQty = Convert.ToInt32(oForm.Items.Item("TotalQty").Specific.Value.ToString().Trim());
				BadQty = Convert.ToInt32(oForm.Items.Item("BadQty").Specific.Value.ToString().Trim());
				BadNote = oForm.Items.Item("BadNote").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode").Specific.Value.ToString().Trim();
				WorkName = oForm.Items.Item("WorkName").Specific.Value.ToString().Trim();
				LastNote = oForm.Items.Item("LastNote").Specific.Value.ToString().Trim();
				Cost01 = Convert.ToDecimal(oForm.Items.Item("Cost01").Specific.Value.ToString().Trim());
				Cost02 = Convert.ToDecimal(oForm.Items.Item("Cost02").Specific.Value.ToString().Trim());
				Cost03 = Convert.ToDecimal(oForm.Items.Item("Cost03").Specific.Value.ToString().Trim());
				Cost04 = Convert.ToDecimal(oForm.Items.Item("Cost04").Specific.Value.ToString().Trim());
				Cost05 = Convert.ToDecimal(oForm.Items.Item("Cost05").Specific.Value.ToString().Trim());
				Cost06 = Convert.ToDecimal(oForm.Items.Item("Cost06").Specific.Value.ToString().Trim());
				Cost07 = Convert.ToDecimal(oForm.Items.Item("Cost07").Specific.Value.ToString().Trim());
				Cost08 = Convert.ToDecimal(oForm.Items.Item("Cost08").Specific.Value.ToString().Trim());
				CostTot = Convert.ToDecimal(oForm.Items.Item("CostTot").Specific.Value.ToString().Trim());
				Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
				UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

				//DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_QM093H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = 1;
				}
				else
				{
					DocEntry = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}

				sQry = " EXEC [PS_QM093_02] ";
				sQry += "'" + DocEntry + "',";
				sQry += "'" + CLTCOD + "',";
				sQry += "'" + ItemCode + "',";
				sQry += "'" + ItemName + "',";
				sQry += "'" + ItemSpec + "',";
				sQry += "'" + DocDate + "',";
				sQry += "'" + TotalQty + "',";
				sQry += "'" + BadQty + "',";
				sQry += "'" + BadNote + "',";
				sQry += "'" + CpCode + "',";
				sQry += "'" + CpName + "',";
				sQry += "'" + WorkCode + "',";
				sQry += "'" + WorkName + "',";
				sQry += "'" + LastNote + "',";
				sQry += "'" + Cost01 + "',";
				sQry += "'" + Cost02 + "',";
				sQry += "'" + Cost03 + "',";
				sQry += "'" + Cost04 + "',";
				sQry += "'" + Cost05 + "',";
				sQry += "'" + Cost06 + "',";
				sQry += "'" + Cost07 + "',";
				sQry += "'" + Cost08 + "',";
				sQry += "'" + CostTot + "',";
				sQry += "'" + Comments + "',";
				sQry += "'" + UserSign + "'";
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
		/// 데이터 UPDATE
		/// </summary>
		/// <returns></returns>
		private bool PS_QM093_UpdateData()
		{
			bool ReturnValue = false;
			int DocEntry;    //관리번호
			string CLTCOD;   //사업장
			string ItemCode; //작번
			string ItemName; //품명
			string ItemSpec; //규격
			string DocDate;  //날짜
			int TotalQty;    //전체수량
			int BadQty;      //불량수량
			string BadNote;  //불량내용
			string CpCode;   //원인공정
			string CpName;   //원인공정명
			string WorkCode; //작업자
			string WorkName; //작업자명
			string LastNote; //최종판정
			decimal Cost01;  //자재
			decimal Cost02;  //가공
			decimal Cost03;  //설계
			decimal Cost04;  //외주
			decimal Cost05;  //분해조립
			decimal Cost06;  //A/S출장
			decimal Cost07;  //운송
			decimal Cost08;  //지체상금
			decimal CostTot; //계
			string Comments; //비고
			string UserSign; //UserSign
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				ItemSpec = oForm.Items.Item("ItemSpec").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				TotalQty = Convert.ToInt32(oForm.Items.Item("TotalQty").Specific.Value.ToString().Trim());
				BadQty = Convert.ToInt32(oForm.Items.Item("BadQty").Specific.Value.ToString().Trim());
				BadNote = oForm.Items.Item("BadNote").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode").Specific.Value.ToString().Trim();
				WorkName = oForm.Items.Item("WorkName").Specific.Value.ToString().Trim();
				LastNote = oForm.Items.Item("LastNote").Specific.Value.ToString().Trim();
				Cost01 = Convert.ToDecimal(oForm.Items.Item("Cost01").Specific.Value.ToString().Trim());
				Cost02 = Convert.ToDecimal(oForm.Items.Item("Cost02").Specific.Value.ToString().Trim());
				Cost03 = Convert.ToDecimal(oForm.Items.Item("Cost03").Specific.Value.ToString().Trim());
				Cost04 = Convert.ToDecimal(oForm.Items.Item("Cost04").Specific.Value.ToString().Trim());
				Cost05 = Convert.ToDecimal(oForm.Items.Item("Cost05").Specific.Value.ToString().Trim());
				Cost06 = Convert.ToDecimal(oForm.Items.Item("Cost06").Specific.Value.ToString().Trim());
				Cost07 = Convert.ToDecimal(oForm.Items.Item("Cost07").Specific.Value.ToString().Trim());
				Cost08 = Convert.ToDecimal(oForm.Items.Item("Cost08").Specific.Value.ToString().Trim());
				CostTot = Convert.ToDecimal(oForm.Items.Item("CostTot").Specific.Value.ToString().Trim());
				Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
				UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

				sQry = " EXEC [PS_QM093_03] ";
				sQry += "'" + DocEntry + "',";
				sQry += "'" + CLTCOD + "',";
				sQry += "'" + ItemCode + "',";
				sQry += "'" + ItemName + "',";
				sQry += "'" + ItemSpec + "',";
				sQry += "'" + DocDate + "',";
				sQry += "'" + TotalQty + "',";
				sQry += "'" + BadQty + "',";
				sQry += "'" + BadNote + "',";
				sQry += "'" + CpCode + "',";
				sQry += "'" + CpName + "',";
				sQry += "'" + WorkCode + "',";
				sQry += "'" + WorkName + "',";
				sQry += "'" + LastNote + "',";
				sQry += "'" + Cost01 + "',";
				sQry += "'" + Cost02 + "',";
				sQry += "'" + Cost03 + "',";
				sQry += "'" + Cost04 + "',";
				sQry += "'" + Cost05 + "',";
				sQry += "'" + Cost06 + "',";
				sQry += "'" + Cost07 + "',";
				sQry += "'" + Cost08 + "',";
				sQry += "'" + CostTot + "',";
				sQry += "'" + Comments + "',";
				sQry += "'" + UserSign + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_QM093_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_QM093_FlushToItemValue(string oUID)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Cost01":
					case "Cost02":
					case "Cost03":
					case "Cost04":
					case "Cost05":
					case "Cost06":
					case "Cost07":
					case "Cost08":
						PS_QM093_CalculateTotalExp();
						break;
					case "ItemCode":
						oDS_PS_QM093H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", ""));
						oDS_PS_QM093H.SetValue("U_ItemSpec", 0, dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", ""));
						break;
					case "CpCode":
						oDS_PS_QM093H.SetValue("U_CpName", 0, dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", ""));
						break;
					case "WorkCode":
						oDS_PS_QM093H.SetValue("U_WorkName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", ""));
						break;
					case "SItemCode":
						oForm.DataSources.UserDataSources.Item("SItemName").Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						oForm.DataSources.UserDataSources.Item("SItemSpec").Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
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
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
							if (PS_QM093_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM093_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM093_LoadCaption();
							PS_QM093_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM093_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM093_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM093_LoadCaption();
							PS_QM093_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						PS_QM093_FormReset();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_QM093_LoadCaption();
						PS_QM093_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_QM093_DeleteData();
							PS_QM093_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM093_LoadCaption();
							PS_QM093_MTX01();
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
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WorkCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SItemCode", "");
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
						if (pVal.Row > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
						{
							oMat.SelectRow(pVal.Row, true, false);
							//DataSource를 이용하여 각 컨트롤에 값을 출력
							oDS_PS_QM093H.SetValue("DocEntry", 0, oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//관리번호
							oDS_PS_QM093H.SetValue("U_CLTCOD", 0, oMat.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//사업장
							oDS_PS_QM093H.SetValue("U_ItemCode", 0, oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작번
							oDS_PS_QM093H.SetValue("U_ItemName", 0, oMat.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//품명
							oDS_PS_QM093H.SetValue("U_ItemSpec", 0, oMat.Columns.Item("ItemSpec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//규격
							oDS_PS_QM093H.SetValue("U_DocDate", 0, oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//날짜
							oDS_PS_QM093H.SetValue("U_TotalQty", 0, oMat.Columns.Item("TotalQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//전체수량
							oDS_PS_QM093H.SetValue("U_BadQty", 0, oMat.Columns.Item("BadQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//불량수량
							oDS_PS_QM093H.SetValue("U_BadNote", 0, oMat.Columns.Item("BadNote").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//불량내용
							oDS_PS_QM093H.SetValue("U_CpCode", 0, oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//원인공정
							oDS_PS_QM093H.SetValue("U_CpName", 0, oMat.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());     //원인공정명
							oDS_PS_QM093H.SetValue("U_WorkCode", 0, oMat.Columns.Item("WorkCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작업자
							oDS_PS_QM093H.SetValue("U_WorkName", 0, oMat.Columns.Item("WorkName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작업자명
							oDS_PS_QM093H.SetValue("U_LastNote", 0, oMat.Columns.Item("LastNote").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//최종판정
							oDS_PS_QM093H.SetValue("U_Cost01", 0, oMat.Columns.Item("Cost01").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//자재
							oDS_PS_QM093H.SetValue("U_Cost02", 0, oMat.Columns.Item("Cost02").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//가공
							oDS_PS_QM093H.SetValue("U_Cost03", 0, oMat.Columns.Item("Cost03").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//설계
							oDS_PS_QM093H.SetValue("U_Cost04", 0, oMat.Columns.Item("Cost04").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//외주
							oDS_PS_QM093H.SetValue("U_Cost05", 0, oMat.Columns.Item("Cost05").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//분해조립
							oDS_PS_QM093H.SetValue("U_Cost06", 0, oMat.Columns.Item("Cost06").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//A/S출장
							oDS_PS_QM093H.SetValue("U_Cost07", 0, oMat.Columns.Item("Cost07").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//운송
							oDS_PS_QM093H.SetValue("U_Cost08", 0, oMat.Columns.Item("Cost08").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//지체상금
							oDS_PS_QM093H.SetValue("U_CostTot", 0, oMat.Columns.Item("CostTot").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//계
							oDS_PS_QM093H.SetValue("U_Comments", 0, oMat.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//비고

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_QM093_LoadCaption();
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
					PS_QM093_FlushToItemValue(pVal.ItemUID);
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
					PS_QM093_FormResize();
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
						}
						else
						{
							PS_QM093_FlushToItemValue(pVal.ItemUID);

							if (pVal.ItemUID == "MSTCOD")
							{
								oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'", "");
							}
							else if (pVal.ItemUID == "SMSTCOD")
							{
								oForm.Items.Item("SMSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim() + "'", "");
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM093H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM093L);
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
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							PS_QM093_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_QM093_LoadCaption();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
							PS_QM093_Add_MatrixRow(oMat.VisualRowCount, false);
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
							oDS_PS_QM093L.RemoveRecord(oDS_PS_QM093L.Size - 1);
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
