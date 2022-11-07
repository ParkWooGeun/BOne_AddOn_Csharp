using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 현장 불합리 개선 등록
	/// </summary>
	internal class PS_QM170 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_QM170B; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM170.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM170_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM170");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM170_CreateItems();
				PS_QM170_ComboBox_Setting();
				PS_QM170_FormResize();
				PS_QM170_LoadCaption();
				PS_QM170_FormReset();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false); 
				oForm.EnableMenu("1282", true);

				oMat.Columns.Item("Check").Visible = false; //매트릭스 선택 필드 숨김
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
		/// PS_QM170_CreateItems
		/// </summary>
		private void PS_QM170_CreateItems()
		{
			try
			{
				oDS_PS_QM170B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//기본정보
				//관리번호
				oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("DocEntry").Specific.DataBind.SetBound(true, "", "DocEntry");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//등록번호
				oForm.DataSources.UserDataSources.Add("RegNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("RegNo").Specific.DataBind.SetBound(true, "", "RegNo");

				//등록일자
				oForm.DataSources.UserDataSources.Add("RegDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RegDate").Specific.DataBind.SetBound(true, "", "RegDate");

				//사번
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

				//소속팀
				oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

				//소속담당
				oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

				//소속반
				oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

				//분야
				oForm.DataSources.UserDataSources.Add("Fields", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Fields").Specific.DataBind.SetBound(true, "", "Fields");

				//등급
				oForm.DataSources.UserDataSources.Add("Grade", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Grade").Specific.DataBind.SetBound(true, "", "Grade");

				//제목
				oForm.DataSources.UserDataSources.Add("Title", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("Title").Specific.DataBind.SetBound(true, "", "Title");

				//내용&비고
				oForm.DataSources.UserDataSources.Add("Comment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("Comment").Specific.DataBind.SetBound(true, "", "Comment");

				//조회정보
				//관리번호
				oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

				//사업장
				oForm.DataSources.UserDataSources.Add("SBPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SBPLID").Specific.DataBind.SetBound(true, "", "SBPLID");

				//등록번호
				oForm.DataSources.UserDataSources.Add("SRegNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SRegNo").Specific.DataBind.SetBound(true, "", "SRegNo");

				//등록기간(시작)
				oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

				//등록기간(종료)
				oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

				//사번
				oForm.DataSources.UserDataSources.Add("SCntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SCntcCode").Specific.DataBind.SetBound(true, "", "SCntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("SCntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SCntcName").Specific.DataBind.SetBound(true, "", "SCntcName");

				//소속팀
				oForm.DataSources.UserDataSources.Add("STeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("STeamCode").Specific.DataBind.SetBound(true, "", "STeamCode");

				//소속담당
				oForm.DataSources.UserDataSources.Add("SRspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SRspCode").Specific.DataBind.SetBound(true, "", "SRspCode");

				//소속반
				oForm.DataSources.UserDataSources.Add("SClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SClsCode").Specific.DataBind.SetBound(true, "", "SClsCode");

				//분야
				oForm.DataSources.UserDataSources.Add("SFields", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SFields").Specific.DataBind.SetBound(true, "", "SFields");

				//등급
				oForm.DataSources.UserDataSources.Add("SGrade", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SGrade").Specific.DataBind.SetBound(true, "", "SGrade");

				//제목
				oForm.DataSources.UserDataSources.Add("STitle", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("STitle").Specific.DataBind.SetBound(true, "", "STitle");

				//내용&비고
				oForm.DataSources.UserDataSources.Add("SComment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("SComment").Specific.DataBind.SetBound(true, "", "SComment");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_QM170_ComboBox_Setting()
		{
			string BPLID;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//기본정보
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분야
				oForm.Items.Item("Fields").Specific.ValidValues.Add("%", "선택");
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q231'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("Fields").Specific, sQry, "", false, false);
				oForm.Items.Item("Fields").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//등급
				oForm.Items.Item("Grade").Specific.ValidValues.Add("%", "선택");
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q232'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("Grade").Specific, sQry, "", false, false);
				oForm.Items.Item("Grade").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//팀
				oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "선택");
				sQry = "    SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '1'";
				sQry += "               AND U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
				oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//담당
				oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "선택");
				sQry = "    SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '2'";
				sQry += "               AND U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
				oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//반
				oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "선택");
				sQry = "    SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '9'";
				sQry += "               AND U_Char3 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
				oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//조회정보
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분야
				oForm.Items.Item("SFields").Specific.ValidValues.Add("%", "전체");
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q231'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("SFields").Specific, sQry, "", false, false);
				oForm.Items.Item("SFields").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//등급
				oForm.Items.Item("SGrade").Specific.ValidValues.Add("%", "전체");
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q232'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("SGrade").Specific, sQry, "", false, false);
				oForm.Items.Item("SGrade").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//매트릭스
				//사업장
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("BPLID"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
				//
				//소속팀
				sQry = "    SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '1'";
				sQry += "               AND U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("TeamCode"), sQry, "", "");

				//소속담당
				sQry = "  SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '2'";
				sQry += "               AND U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("RspCode"), sQry, "", "");

				//소속반
				sQry = "    SELECT     U_Code,";
				sQry += "               U_CodeNm";
				sQry += " FROM      [@PS_HR200L]";
				sQry += " WHERE     Code = '9'";
				sQry += "               AND U_Char3 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim()  + "'";
				sQry += "               AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ClsCode"), sQry, "", "");

				//분야
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q231'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Fields"), sQry, "", "");

				//등급
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName As [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q232'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Grade"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM170_FormResize
		/// </summary>
		private void PS_QM170_FormResize()
		{
			try
			{
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM170_LoadCaption
		/// </summary>
		private void PS_QM170_LoadCaption()
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
		/// 화면 초기화
		/// </summary>
		private void PS_QM170_FormReset()
		{
			string User_BPLId;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				//관리번호
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [Z_PS_QM170_01]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oForm.DataSources.UserDataSources.Item("DocEntry").Value = "1";
				}
				else
				{
					oForm.DataSources.UserDataSources.Item("DocEntry").Value = Convert.ToString(Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
				}

				User_BPLId = dataHelpClass.User_BPLID();

				//기준정보
				oForm.DataSources.UserDataSources.Item("BPLID").Value = User_BPLId;	//사업장
				oForm.DataSources.UserDataSources.Item("RegNo").Value = "";		//등록번호
				oForm.DataSources.UserDataSources.Item("RegDate").Value = DateTime.Now.ToString("yyyyMMdd"); //등록일자
				oForm.DataSources.UserDataSources.Item("CntcCode").Value = "";	//사번
				oForm.DataSources.UserDataSources.Item("Fields").Value = "%";	//분야
				oForm.DataSources.UserDataSources.Item("Grade").Value = "%";	//등급
				oForm.DataSources.UserDataSources.Item("TeamCode").Value = "%";	//팀
				oForm.DataSources.UserDataSources.Item("RspCode").Value = "%";	//담당
				oForm.DataSources.UserDataSources.Item("ClsCode").Value = "%";	//반
				oForm.DataSources.UserDataSources.Item("Title").Value = "";		//제목
				oForm.DataSources.UserDataSources.Item("Comment").Value = "";	//내용&비고
				PS_QM170_GetRegNo(); //등록번호

				//조회정보
				oForm.DataSources.UserDataSources.Item("SDocEntry").Value = "";	//관리번호
				oForm.Items.Item("SBPLID").Specific.Select(User_BPLId, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.DataSources.UserDataSources.Item("SRegNo").Value = "";	//등록번호
				oForm.Items.Item("SFrDate").Specific.Value = "";				//Format(Now, "yyyyMM01") '등록기간(시작)
				oForm.Items.Item("SToDate").Specific.Value = "";				//Format(Now, "yyyyMMdd") '등록기간(종료)
				oForm.DataSources.UserDataSources.Item("SCntcCode").Value = "";	//사번
				oForm.DataSources.UserDataSources.Item("SFields").Value = "%";	//분야
				oForm.DataSources.UserDataSources.Item("SGrade").Value = "%";	//등급
				oForm.DataSources.UserDataSources.Item("STeamCode").Value = "%";//팀
				oForm.DataSources.UserDataSources.Item("SRspCode").Value = "%";	//담당
				oForm.DataSources.UserDataSources.Item("SClsCode").Value = "%";	//반
				oForm.DataSources.UserDataSources.Item("STitle").Value = "";	//제목
				oForm.DataSources.UserDataSources.Item("SComment").Value = "";	//내용&비고

				oForm.Items.Item("CntcCode").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 등록번호 생성
		/// </summary>
		private void PS_QM170_GetRegNo()
		{
			string RegDate;
			string BPLID;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				RegDate = oForm.Items.Item("RegDate").Specific.Value.ToString().Trim().Substring(0, 6);

				sQry = "EXEC PS_QM170_05 '" + BPLID + "', '" + RegDate + "'";
				oRecordSet.DoQuery(sQry);

				oForm.DataSources.UserDataSources.Item("RegNo").Value = oRecordSet.Fields.Item("RegNo").Value.ToString().Trim();
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
		/// PS_QM170_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM170_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_QM170B.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_QM170B.Offset = oRow;
				oDS_PS_QM170B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		private void PS_QM170_MTX01()
		{
			int i;
			string sDocEntry; //관리번호(PK)
			string SBPLID;    //사업장코드(PK)
			string SRegNo;    //등록번호(PK)
			string SFrDate;   //등록일자(From)
			string SToDate;   //등록일자(To)
			string SCntcCode; //개선자사번
			string STeamCode; //소속팀코드
			string SRspCode;  //소속담당코드
			string SClsCode;  //소속반코드
			string SFields;   //분야
			string SGrade;    //등급
			string STitle;    //제목(%)
			string SComment;  //비고(%)
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				sDocEntry = oForm.DataSources.UserDataSources.Item("SDocEntry").Value.ToString().Trim();
				SBPLID = oForm.DataSources.UserDataSources.Item("SBPLID").Value.ToString().Trim();
				SRegNo = oForm.DataSources.UserDataSources.Item("SRegNo").Value.ToString().Trim();
				SFrDate = oForm.DataSources.UserDataSources.Item("SFrDate").Value.ToString().Trim();
				SToDate = oForm.DataSources.UserDataSources.Item("SToDate").Value.ToString().Trim();
				SCntcCode = oForm.DataSources.UserDataSources.Item("SCntcCode").Value.ToString().Trim();
				STeamCode = oForm.DataSources.UserDataSources.Item("STeamCode").Value.ToString().Trim();
				SRspCode = oForm.DataSources.UserDataSources.Item("SRspCode").Value.ToString().Trim();
				SClsCode = oForm.DataSources.UserDataSources.Item("SClsCode").Value.ToString().Trim();
				SFields = oForm.DataSources.UserDataSources.Item("SFields").Value.ToString().Trim();
				SGrade = oForm.DataSources.UserDataSources.Item("SGrade").Value.ToString().Trim();
				STitle = oForm.DataSources.UserDataSources.Item("STitle").Value.ToString().Trim();
				SComment = oForm.DataSources.UserDataSources.Item("SComment").Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "  EXEC [PS_QM170_01] '";
				sQry += sDocEntry + "','";
				sQry += SBPLID + "','";
				sQry += SRegNo + "','";
				sQry += SFrDate + "','";
				sQry += SToDate + "','";
				sQry += SCntcCode + "','";
				sQry += STeamCode + "','";
				sQry += SRspCode + "','";
				sQry += SClsCode + "','";
				sQry += SFields + "','";
				sQry += SGrade + "','";
				sQry += STitle + "','";
				sQry += SComment + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM170B.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_QM170_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM170B.Size)
					{
						oDS_PS_QM170B.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_QM170B.Offset = i;

					oDS_PS_QM170B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM170B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("BPLID").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("RegNo").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColDt01", i, oRecordSet.Fields.Item("RegDate").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("ClsCode").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("Fields").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("Grade").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("Title").Value.ToString().Trim());
					oDS_PS_QM170B.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());
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
		private void PS_QM170_DeleteData()
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

					sQry = "SELECT COUNT(*) FROM [Z_PS_QM170_01] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "EXEC PS_QM170_04 '" + DocEntry + "'";
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
		private bool PS_QM170_UpdateData()
		{
			bool ReturnValue = false;
			int DocEntry;	 //관리번호(PK)
			string BPLID;	 //사업장코드(PK)
			string RegNo;	 //등록번호(PK)
			string RegDate;	 //등록일자
			string CntcCode; //개선자사번
			string CntcName; //개선자성명
			string TeamCode; //소속팀코드
			string RspCode;	 //소속담당코드
			string ClsCode;	 //소속반코드
			string Fields;	 //분야
			string Grade;	 //등급
			string Title;	 //제목
			string Comment;  //비고
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				RegNo = oForm.Items.Item("RegNo").Specific.Value.ToString().Trim();
				RegDate = oForm.Items.Item("RegDate").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				Fields = oForm.Items.Item("Fields").Specific.Value.ToString().Trim();
				Grade = oForm.Items.Item("Grade").Specific.Value.ToString().Trim();
				Title = oForm.Items.Item("Title").Specific.Value.ToString().Trim();
				Comment = oForm.Items.Item("Comment").Specific.Value.ToString().Trim();

				sQry = " EXEC [PS_QM170_03] '";
				sQry += DocEntry + "','";
				sQry += BPLID + "','";
				sQry += RegNo + "','";
				sQry += RegDate + "','";
				sQry += CntcCode + "','";
				sQry += CntcName + "','";
				sQry += TeamCode + "','";
				sQry += RspCode + "','";
				sQry += ClsCode + "','";
				sQry += Fields + "','";
				sQry += Grade + "','";
				sQry += Title + "','";
				sQry += Comment + "'";
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
		/// 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_QM170_AddData()
		{
			bool ReturnValue = false;
			int DocEntry;    //관리번호(PK)
			string BPLID;    //사업장코드(PK)
			string RegNo;    //등록번호(PK)
			string RegDate;  //등록일자
			string CntcCode; //개선자사번
			string CntcName; //개선자성명
			string TeamCode; //소속팀코드
			string RspCode;  //소속담당코드
			string ClsCode;  //소속반코드
			string Fields;   //분야
			string Grade;    //등급
			string Title;    //제목
			string Comment;  //비고
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				RegNo = oForm.Items.Item("RegNo").Specific.Value.ToString().Trim();
				RegDate = oForm.Items.Item("RegDate").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				Fields = oForm.Items.Item("Fields").Specific.Value.ToString().Trim();
				Grade = oForm.Items.Item("Grade").Specific.Value.ToString().Trim();
				Title = oForm.Items.Item("Title").Specific.Value.ToString().Trim();
				Comment = oForm.Items.Item("Comment").Specific.Value.ToString().Trim();

				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[Z_PS_QM170_01]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = 1;
				}
				else
				{
					DocEntry = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}

				sQry = "  EXEC [PS_QM170_02] '";
				sQry += DocEntry + "','";
				sQry += BPLID + "','";
				sQry += RegNo + "','";
				sQry += RegDate + "','";
				sQry += CntcCode + "','";
				sQry += CntcName + "','";
				sQry += TeamCode + "','";
				sQry += RspCode + "','";
				sQry += ClsCode + "','";
				sQry += Fields + "','";
				sQry += Grade + "','";
				sQry += Title + "','";
				sQry += Comment + "'";
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
		private bool PS_QM170_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("RegNo").Specific.Value.ToString().Trim()))
				{
					errMessage = "등록번호는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("RegDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "등록일자는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "사원번호는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("Fields").Specific.Value.ToString().Trim() == "%")
				{
					errMessage = "분야는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("Grade").Specific.Value.ToString().Trim() == "%")
				{
					errMessage = "등급은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("Title").Specific.Value.ToString().Trim()))
				{
					errMessage = "제목은 필수사항입니다. 확인하세요.";
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
		/// PS_QM170_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM170_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int loopCount;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "SBPLID":
						if (oForm.Items.Item("STeamCode").Specific.ValidValues.Count > 0)
						{
							for (loopCount = oForm.Items.Item("STeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
							{
								oForm.Items.Item("STeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}
						oForm.Items.Item("STeamCode").Specific.ValidValues.Add("%", "전체");
						sQry = "    SELECT     U_Code,";
						sQry += "               U_CodeNm";
						sQry += " FROM      [@PS_HR200L]";
						sQry += " WHERE     Code = '1'";
						sQry += "               AND U_Char2 = '" + oForm.Items.Item("SBPLID").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_UseYN = 'Y'";
						sQry += " ORDER BY U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("STeamCode").Specific, sQry, "", false, false);
						oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;
					case "STeamCode":
						if (oForm.Items.Item("SRspCode").Specific.ValidValues.Count > 0)
						{
							for (loopCount = oForm.Items.Item("SRspCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
							{
								oForm.Items.Item("SRspCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}
						oForm.Items.Item("SRspCode").Specific.ValidValues.Add("%", "전체");
						sQry = "    SELECT     U_Code,";
						sQry += "               U_CodeNm";
						sQry += " FROM      [@PS_HR200L]";
						sQry += " WHERE     Code = '2'";
						sQry += "               AND U_Char2 = '" + oForm.Items.Item("SBPLID").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_Char1 = '" + oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_UseYN = 'Y'";
						sQry += " ORDER BY U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("SRspCode").Specific, sQry, "", false, false);
						oForm.Items.Item("SRspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;
					case "SRspCode":
						if (oForm.Items.Item("SClsCode").Specific.ValidValues.Count > 0)
						{
							for (loopCount = oForm.Items.Item("SClsCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
							{
								oForm.Items.Item("SClsCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}
						oForm.Items.Item("SClsCode").Specific.ValidValues.Add("%", "전체");
						sQry = "  SELECT     U_Code, ";
						sQry += "            U_CodeNm ";
						sQry += " FROM    [@PS_HR200L] ";
						sQry += " WHERE     Code = '9' ";
						sQry += "               AND U_Char3 = '" + oForm.Items.Item("SBPLID").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_Char2 = '" + oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_Char1 = '" + oForm.Items.Item("SRspCode").Specific.Value.ToString().Trim() + "'";
						sQry += "               AND U_UseYN = 'Y'";
						sQry += " ORDER BY U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("SClsCode").Specific, sQry, "", false, false);
						oForm.Items.Item("SClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;
					case "CntcCode":
						oForm.DataSources.UserDataSources.Item("CntcName").Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.DataSources.UserDataSources.Item("TeamCode").Value = dataHelpClass.Get_ReData("CASE WHEN ISNULL(U_TeamCode, '') = '' THEN '%' ELSE U_TeamCode END", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.DataSources.UserDataSources.Item("RspCode").Value = dataHelpClass.Get_ReData("CASE WHEN ISNULL(U_RspCode, '') = '' THEN '%' ELSE U_RspCode END", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.DataSources.UserDataSources.Item("ClsCode").Value = dataHelpClass.Get_ReData("CASE WHEN ISNULL(U_ClsCode, '') = '' THEN '%' ELSE U_ClsCode END", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
						break;
					case "SCntcCode":
						oForm.Items.Item("SCntcName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("SCntcCode").Specific.Value.ToString().Trim() + "'", "");
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
                //	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8	
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
							if (PS_QM170_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_QM170_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM170_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM170_LoadCaption();
							PS_QM170_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM170_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_QM170_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM170_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM170_LoadCaption();
							PS_QM170_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_QM170_LoadCaption();
						PS_QM170_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_QM170_DeleteData();
							PS_QM170_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM170_LoadCaption();
							PS_QM170_MTX01();
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
					PS_QM170_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
							oForm.DataSources.UserDataSources.Item("DocEntry").Value = oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("BPLID").Value = oMat.Columns.Item("BPLID").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("RegNo").Value = oMat.Columns.Item("RegNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("RegDate").Value = oMat.Columns.Item("RegDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("CntcCode").Value = oMat.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("CntcName").Value = oMat.Columns.Item("CntcName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("TeamCode").Value = oMat.Columns.Item("TeamCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("RspCode").Value = oMat.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("ClsCode").Value = oMat.Columns.Item("ClsCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("Fields").Value = oMat.Columns.Item("Fields").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("Grade").Value = oMat.Columns.Item("Grade").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("Title").Value = oMat.Columns.Item("Title").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("Comment").Value = oMat.Columns.Item("Comment").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_QM170_LoadCaption();
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
							PS_QM170_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_QM170_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM170B);
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
						case "1283": //삭제
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
							PS_QM170_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_QM170_LoadCaption();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
							PS_QM170_Add_MatrixRow(oMat.VisualRowCount, false);
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
							oDS_PS_QM170B.RemoveRecord(oDS_PS_QM170B.Size - 1);
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
