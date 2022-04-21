using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 전산장비조회
	/// </summary>
	internal class PS_GA167 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.Grid oGrid04;
		private SAPbouiCOM.Grid oGrid05;

		private SAPbouiCOM.DataTable oDS_PS_GA167A;
		private SAPbouiCOM.DataTable oDS_PS_GA167B;
		private SAPbouiCOM.DataTable oDS_PS_GA167C;
		private SAPbouiCOM.DataTable oDS_PS_GA167D;
		private SAPbouiCOM.DataTable oDS_PS_GA167E;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_GA167.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_GA167_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_GA167");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_GA167_CreateItems();
				PS_GA167_ComboBox_Setting();
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
		/// PS_GA167_CreateItems
		/// </summary>
		private void PS_GA167_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;
				oGrid04 = oForm.Items.Item("Grid04").Specific;
				oGrid05 = oForm.Items.Item("Grid05").Specific;

				oForm.DataSources.DataTables.Add("PS_GA167A");
				oForm.DataSources.DataTables.Add("PS_GA167B");
				oForm.DataSources.DataTables.Add("PS_GA167C");
				oForm.DataSources.DataTables.Add("PS_GA167D");
				oForm.DataSources.DataTables.Add("PS_GA167E");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_GA167A");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_GA167B");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_GA167C");
				oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_GA167D");
				oGrid05.DataTable = oForm.DataSources.DataTables.Item("PS_GA167E");

				oDS_PS_GA167A = oForm.DataSources.DataTables.Item("PS_GA167A");
				oDS_PS_GA167B = oForm.DataSources.DataTables.Item("PS_GA167B");
				oDS_PS_GA167C = oForm.DataSources.DataTables.Item("PS_GA167C");
				oDS_PS_GA167D = oForm.DataSources.DataTables.Item("PS_GA167D");
				oDS_PS_GA167E = oForm.DataSources.DataTables.Item("PS_GA167E");


				//장비리스트
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//팀코드
				oForm.DataSources.UserDataSources.Add("TeamCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("TeamCode01").Specific.DataBind.SetBound(true, "", "TeamCode01");

				//팀명
				oForm.DataSources.UserDataSources.Add("TeamName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("TeamName01").Specific.DataBind.SetBound(true, "", "TeamName01");

				//사용자사번
				oForm.DataSources.UserDataSources.Add("CntcCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode01").Specific.DataBind.SetBound(true, "", "CntcCode01");

				//사용자성명
				oForm.DataSources.UserDataSources.Add("CntcName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName01").Specific.DataBind.SetBound(true, "", "CntcName01");

				//장비분류
				oForm.DataSources.UserDataSources.Add("Ctgr01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("Ctgr01").Specific.DataBind.SetBound(true, "", "Ctgr01");

				//모델명
				oForm.DataSources.UserDataSources.Add("ModelNm01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ModelNm01").Specific.DataBind.SetBound(true, "", "ModelNm01");

				//위치구분
				oForm.DataSources.UserDataSources.Add("LocCls01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("LocCls01").Specific.DataBind.SetBound(true, "", "LocCls01");

				//관리번호
				oForm.DataSources.UserDataSources.Add("MngNo01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("MngNo01").Specific.DataBind.SetBound(true, "", "MngNo01");

				//등록기간(시작)
				oForm.DataSources.UserDataSources.Add("RegFrDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RegFrDt01").Specific.DataBind.SetBound(true, "", "RegFrDt01");

				//등록기간(종료)
				oForm.DataSources.UserDataSources.Add("RegToDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("RegToDt01").Specific.DataBind.SetBound(true, "", "RegToDt01");

				//폐기장비 포함
				oForm.DataSources.UserDataSources.Add("DisUseYN01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("DisUseYN01").Specific.DataBind.SetBound(true, "", "DisUseYN01");
				oForm.Items.Item("DisUseYN01").Specific.Checked = false;

				//장비별사용자이력
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//사용자사번
				oForm.DataSources.UserDataSources.Add("CntcCode02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode02").Specific.DataBind.SetBound(true, "", "CntcCode02");

				//사용자성명
				oForm.DataSources.UserDataSources.Add("CntcName02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName02").Specific.DataBind.SetBound(true, "", "CntcName02");

				//장비분류
				oForm.DataSources.UserDataSources.Add("Ctgr02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("Ctgr02").Specific.DataBind.SetBound(true, "", "Ctgr02");

				//모델명
				oForm.DataSources.UserDataSources.Add("ModelNm02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ModelNm02").Specific.DataBind.SetBound(true, "", "ModelNm02");

				//사용기간(시작)
				oForm.DataSources.UserDataSources.Add("UseFrDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("UseFrDt02").Specific.DataBind.SetBound(true, "", "UseFrDt02");

				//사용기간(종료)
				oForm.DataSources.UserDataSources.Add("UseToDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("UseToDt02").Specific.DataBind.SetBound(true, "", "UseToDt02");

				//관리번호
				oForm.DataSources.UserDataSources.Add("MngNo02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("MngNo02").Specific.DataBind.SetBound(true, "", "MngNo02");

				//주용도(%)
				oForm.DataSources.UserDataSources.Add("MainUse02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("MainUse02").Specific.DataBind.SetBound(true, "", "MainUse02");

				//사용자별장비이력
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID03").Specific.DataBind.SetBound(true, "", "BPLID03");

				//사용자사번
				oForm.DataSources.UserDataSources.Add("CntcCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode03").Specific.DataBind.SetBound(true, "", "CntcCode03");

				//사용기간(시작)
				oForm.DataSources.UserDataSources.Add("UseFrDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("UseFrDt03").Specific.DataBind.SetBound(true, "", "UseFrDt03");

				//사용기간(종료)
				oForm.DataSources.UserDataSources.Add("UseToDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("UseToDt03").Specific.DataBind.SetBound(true, "", "UseToDt03");

				//장비분류
				oForm.DataSources.UserDataSources.Add("Ctgr03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("Ctgr03").Specific.DataBind.SetBound(true, "", "Ctgr03");

				//장비별 점검이력
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLID05").Specific.DataBind.SetBound(true, "", "BPLID05");

				//사용자사번
				oForm.DataSources.UserDataSources.Add("CntcCode05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode05").Specific.DataBind.SetBound(true, "", "CntcCode05");

				//사용자성명
				oForm.DataSources.UserDataSources.Add("CntcName05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName05").Specific.DataBind.SetBound(true, "", "CntcName05");

				//장비분류
				oForm.DataSources.UserDataSources.Add("Ctgr05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("Ctgr05").Specific.DataBind.SetBound(true, "", "Ctgr05");

				//모델명
				oForm.DataSources.UserDataSources.Add("ModelNm05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ModelNm05").Specific.DataBind.SetBound(true, "", "ModelNm05");

				//위치구분
				oForm.DataSources.UserDataSources.Add("LocCls05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("LocCls05").Specific.DataBind.SetBound(true, "", "LocCls05");

				//관리번호
				oForm.DataSources.UserDataSources.Add("MngNo05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("MngNo05").Specific.DataBind.SetBound(true, "", "MngNo05");

				//점검기간(시작)
				oForm.DataSources.UserDataSources.Add("ChkFrDt05", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ChkFrDt05").Specific.DataBind.SetBound(true, "", "ChkFrDt05");

				//점검기간(종료)
				oForm.DataSources.UserDataSources.Add("ChkToDt05", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ChkToDt05").Specific.DataBind.SetBound(true, "", "ChkToDt05");

				//점검분류
				oForm.DataSources.UserDataSources.Add("ChkCls05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ChkCls05").Specific.DataBind.SetBound(true, "", "ChkCls05");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_ComboBox_Setting
		/// </summary>
		private void PS_GA167_ComboBox_Setting()
		{
			string sQry;
			string BPLID;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//장비리스트
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분류
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '12'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("Ctgr01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("Ctgr01").Specific, sQry, "%", false, false);

				//위치구분
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '14'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("LocCls01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("LocCls01").Specific, sQry, "%", false, false);

				//장비별사용자이력
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID03").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분류
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '12'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("Ctgr03").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("Ctgr03").Specific, sQry, "%", false, false);

				//장비별사용자이력
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분류
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '12'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("Ctgr02").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("Ctgr02").Specific, sQry, "%", false, false);

				//장비 점검이력
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID05").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);

				//분류
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '12'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("Ctgr05").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("Ctgr05").Specific, sQry, "%", false, false);

				//위치구분
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '14'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("LocCls05").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("LocCls05").Specific, sQry, "%", false, false);

				//점검분류
				sQry = " SELECT   U_Code,";
				sQry += "          U_CodeNm";
				sQry += " FROM     [@PS_GA050L]";
				sQry += " WHERE    Code = '18'";
				sQry += "          AND U_UseYN = 'Y'";
				sQry += " ORDER BY U_Seq";
				oForm.Items.Item("ChkCls05").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ChkCls05").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_GA167_FlushToItemValue(string oUID)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "CntcCode01":
						if (oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim() == "9999999")
						{
							oForm.Items.Item("CntcName01").Specific.Value = "공용"; //성명
						}
						else
						{
							oForm.Items.Item("CntcName01").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim() + "'", ""); //성명
						}
						break;

					case "TeamCode01":
						if (oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim() == oForm.Items.Item("BPLID01").Specific.Selected.Value + "999")
						{
							oForm.Items.Item("TeamName01").Specific.Value = "전체공용";
						}
						else if (oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim() == "Z" + oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim() + "99")
						{
							oForm.Items.Item("TeamName01").Specific.Value = "사용부서없음";
						}
						else
						{
							oForm.Items.Item("TeamName01").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim() + "'", " AND Code = '1'"); //팀
						}
						break;

					case "CntcCode02":
						if (oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim() == "9999999")
						{
							oForm.Items.Item("CntcName02").Specific.Value = "공용"; //성명
						}
						else
						{
							oForm.Items.Item("CntcName02").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim() + "'", ""); //성명
						}
						break;

					case "CntcCode03":
						if (oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim() == "9999999")
						{
							oForm.Items.Item("CntcName03").Specific.Value = "공용"; //성명
						}
						else
						{
							oForm.Items.Item("CntcName03").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim() + "'", ""); //성명
						}
						break;

					case "CntcCode05":
						if (oForm.Items.Item("CntcCode05").Specific.Value.ToString().Trim() == "9999999")
						{
							oForm.Items.Item("CntcName05").Specific.Value = "공용"; //성명
						}
						else
						{
							oForm.Items.Item("CntcName05").Specific.Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("CntcCode05").Specific.Value.ToString().Trim() + "'", ""); //성명
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_GetDetail
		/// </summary>
		private void PS_GA167_GetDetail()
		{
			int loopCount1;
			string MngNo = string.Empty;

			try
			{
				for (loopCount1 = 0; loopCount1 <= oGrid01.Rows.Count - 1; loopCount1++)
				{
					if (oGrid01.Rows.IsSelected(loopCount1) == true)
					{
						MngNo = oGrid01.DataTable.GetValue(5, loopCount1);
					}
				}

				PS_GA160 oTempClass = new PS_GA160();
				oTempClass.LoadForm(MngNo);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_MTX01
		/// </summary>
		private void PS_GA167_MTX01()
		{
			string BPLID;    //사업장
			string TeamCode; //팀코드
			string CntcCode; //사용자
			string Ctgr;     //전산장비분류
			string ModelNm;  //모델명
			string LocCls;   //위치구분
			string MngNo;    //관리번호
			string RegFrDt;  //등록기간(시작)
			string RegToDt;  //등록기간(종료)
			string DisUseYN; //폐기장비 포함 여부
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr01").Specific.Selected.Value.ToString().Trim();
				ModelNm = oForm.Items.Item("ModelNm01").Specific.Value.ToString().Trim();
				LocCls = oForm.Items.Item("LocCls01").Specific.Selected.Value.ToString().Trim();
				MngNo = oForm.Items.Item("MngNo01").Specific.Value.ToString().Trim();
				RegFrDt = oForm.Items.Item("RegFrDt01").Specific.Value.ToString().Trim();
				RegToDt = oForm.Items.Item("RegToDt01").Specific.Value.ToString().Trim();
				if (oForm.DataSources.UserDataSources.Item("DisUseYN01").Value == "Y")
				{
					DisUseYN = "Y";
				}
				else
				{
					DisUseYN = "N";
				}

				ProgressBar01.Text = "조회시작!";

				if (Ctgr == "M") //모니터로 조회할 경우는 모니터만 단독으로 출력
				{
					sQry = " EXEC PS_GA167_06 '";
					sQry += BPLID + "','";
					sQry += TeamCode + "','";
					sQry += CntcCode + "','";
					sQry += Ctgr + "','";
					sQry += ModelNm + "','";
					sQry += LocCls + "','";
					sQry += MngNo + "','";
					sQry += RegFrDt + "','";
					sQry += RegToDt + "','";
					sQry += DisUseYN + "'";
				}
				else
				{
					sQry = " EXEC PS_GA167_01 '";
					sQry += BPLID + "','";
					sQry += TeamCode + "','";
					sQry += CntcCode + "','";
					sQry += Ctgr + "','";
					sQry += ModelNm + "','";
					sQry += LocCls + "','";
					sQry += MngNo + "','";
					sQry += RegFrDt + "','";
					sQry += RegToDt + "','";
					sQry += DisUseYN + "'";
				}

				oGrid01.DataTable.Clear();
				oDS_PS_GA167A.ExecuteQuery(sQry);

				if (Ctgr == "M")
				{
					oGrid01.Columns.Item(12).RightJustified = true;
				}
				else
				{
					oGrid01.Columns.Item(18).RightJustified = true;
				}

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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_GA167_MTX02
		/// </summary>
		private void PS_GA167_MTX02()
		{
			string BPLID;	 //사업장
			string CntcCode; //사용자
			string Ctgr;	 //전산장비분류
			string ModelNm;	 //모델명
			string UseFrDt;	 //사용기간(시작)
			string UseToDt;	 //사용기간(종료)
			string MngNo;	 //관리번호
			string MainUse;  //주용도
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID02").Specific.Selected.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode02").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr02").Specific.Selected.Value.ToString().Trim();
				ModelNm = oForm.Items.Item("ModelNm02").Specific.Value.ToString().Trim();
				UseFrDt = oForm.Items.Item("UseFrDt02").Specific.Value.ToString().Trim();
				UseToDt = oForm.Items.Item("UseToDt02").Specific.Value.ToString().Trim();
				MngNo = oForm.Items.Item("MngNo02").Specific.Value.ToString().Trim();
				MainUse = oForm.Items.Item("MainUse02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = " EXEC PS_GA167_02 '";
				sQry += BPLID + "','";
				sQry += CntcCode + "','";
				sQry += Ctgr + "','";
				sQry += ModelNm + "','";
				sQry += UseFrDt + "','";
				sQry += UseToDt + "','";
				sQry += MngNo + "','";
				sQry += MainUse + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_GA167B.ExecuteQuery(sQry);

				if (oGrid02.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid02.AutoResizeColumns();
				oForm.Update();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_GA167_MTX03
		/// </summary>
		private void PS_GA167_MTX03()
		{
			string BPLID;	 //사업장
			string CntcCode; //사용자
			string UseFrDt;	 //사용기간(시작)
			string UseToDt;	 //사용기간(종료)
			string Ctgr;     //전산장비분류
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID03").Specific.Selected.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode03").Specific.Value.ToString().Trim();
				UseFrDt = oForm.Items.Item("UseFrDt03").Specific.Value.ToString().Trim();
				UseToDt = oForm.Items.Item("UseToDt03").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr03").Specific.Selected.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = " EXEC PS_GA167_03 '";
				sQry += BPLID + "','";
				sQry += CntcCode + "','";
				sQry += UseFrDt + "','";
				sQry += UseToDt + "','";
				sQry += Ctgr + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_GA167C.ExecuteQuery(sQry);
				
				if (oGrid03.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid03.AutoResizeColumns();
				oForm.Update();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_GA167_MTX05
		/// </summary>
		private void PS_GA167_MTX05()
		{
			string BPLID;	 //사업장
			string CntcCode; //사용자
			string Ctgr;	 //전산장비분류
			string ModelNm;	 //모델명
			string LocCls;	 //위치구분
			string MngNo;	 //관리번호
			string ChkFrDt;	 //점검기간(시작)
			string ChkToDt;	 //점검기간(종료)
			string ChkCls;   //점검분류
			string sQry;
			string errMessage = string.Empty;
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{

				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID05").Specific.Selected.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode05").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr05").Specific.Selected.Value.ToString().Trim();
				ModelNm = oForm.Items.Item("ModelNm05").Specific.Value.ToString().Trim();
				LocCls = oForm.Items.Item("LocCls05").Specific.Selected.Value.ToString().Trim();
				MngNo = oForm.Items.Item("MngNo05").Specific.Value.ToString().Trim();
				ChkFrDt = oForm.Items.Item("ChkFrDt05").Specific.Value.ToString().Trim();
				ChkToDt = oForm.Items.Item("ChkToDt05").Specific.Value.ToString().Trim();
				ChkCls = oForm.Items.Item("ChkCls05").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = " EXEC PS_GA167_05 '";
				sQry += BPLID + "','";
				sQry += CntcCode + "','";
				sQry += Ctgr + "','";
				sQry += ModelNm + "','";
				sQry += LocCls + "','";
				sQry += MngNo + "','";
				sQry += ChkFrDt + "','";
				sQry += ChkToDt + "','";
				sQry += ChkCls + "'";

				oGrid05.DataTable.Clear();
				oDS_PS_GA167E.ExecuteQuery(sQry);
			
				oGrid05.Columns.Item(10).RightJustified = true;
				oGrid05.Columns.Item(19).RightJustified = true;

				if (oGrid05.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid05.AutoResizeColumns();
				oForm.Update();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_GA167_FormResize
		/// </summary>
		private void PS_GA167_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 140;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				if (oGrid01.Columns.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}
				if (oGrid02.Columns.Count > 0)
				{
					oGrid02.AutoResizeColumns();
				}
				if (oGrid03.Columns.Count > 0)
				{
					oGrid03.AutoResizeColumns();
				}
				if (oGrid04.Columns.Count > 0)
				{
					oGrid04.AutoResizeColumns();
				}
				if (oGrid05.Columns.Count > 0)
				{
					oGrid05.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_GA167_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLID;	 //사업장
			string TeamCode; //팀(부서)
			string CntcCode; //사용자
			string Ctgr;	 //전산장비분류
			string ModelNm;	 //모델명
			string LocCls;	 //위치구분
			string MngNo;	 //관리번호
			string RegFrDt;	 //등록기간(시작)
			string RegToDt;	 //등록기간(종료)
			string DisUseYN; //폐기장비 포함 여부
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID01").Specific.Selected.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode01").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr01").Specific.Selected.Value.ToString().Trim();
				ModelNm = oForm.Items.Item("ModelNm01").Specific.Value.ToString().Trim();
				LocCls = oForm.Items.Item("LocCls01").Specific.Selected.Value.ToString().Trim();
				MngNo = oForm.Items.Item("MngNo01").Specific.Value.ToString().Trim();
				RegFrDt = oForm.Items.Item("RegFrDt01").Specific.Value.ToString().Trim();
				RegToDt = oForm.Items.Item("RegToDt01").Specific.Value.ToString().Trim();
				if (oForm.DataSources.UserDataSources.Item("DisUseYN01").Value == "Y")
				{
					DisUseYN = "Y";
				}
				else
				{
					DisUseYN = "N";
				}

				if (string.IsNullOrEmpty(RegFrDt))
                {
					RegFrDt = "19000101";
				}
				if (string.IsNullOrEmpty(RegToDt))
				{
					RegToDt = "21000101";
				}

				WinTitle = "[PS_GA167] 레포트";
				ReportName = "PS_GA167_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@Ctgr", Ctgr));
				dataPackParameter.Add(new PSH_DataPackClass("@ModelNm", ModelNm));
				dataPackParameter.Add(new PSH_DataPackClass("@LocCls", LocCls));
				dataPackParameter.Add(new PSH_DataPackClass("@MngNo", MngNo));
				dataPackParameter.Add(new PSH_DataPackClass("@RegFrDt", DateTime.ParseExact(RegFrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@RegToDt", DateTime.ParseExact(RegToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@DisUseYN", DisUseYN));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA167_Print_Report05
		/// </summary>
		[STAThread]
		private void PS_GA167_Print_Report05()
		{
			string WinTitle;
			string ReportName;
			string BPLID;	 //사업장
			string CntcCode; //사용자
			string Ctgr;	 //전산장비분류
			string ModelNm;	 //모델명
			string LocCls;	 //위치구분
			string MngNo;	 //관리번호
			string ChkFrDt;	 //점검기간(시작)
			string ChkToDt;	 //점검기간(종료)
			string ChkCls;   //점검분류
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID05").Specific.Selected.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode05").Specific.Value.ToString().Trim();
				Ctgr = oForm.Items.Item("Ctgr05").Specific.Selected.Value.ToString().Trim();
				ModelNm = oForm.Items.Item("ModelNm05").Specific.Value.ToString().Trim();
				LocCls = oForm.Items.Item("LocCls05").Specific.Selected.Value.ToString().Trim();
				MngNo = oForm.Items.Item("MngNo05").Specific.Value.ToString().Trim();
				ChkFrDt = oForm.Items.Item("ChkFrDt05").Specific.Value.ToString().Trim();
				ChkToDt = oForm.Items.Item("ChkToDt05").Specific.Value.ToString().Trim();
				ChkCls = oForm.Items.Item("ChkCls05").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(ChkFrDt))
				{
					ChkFrDt = "19000101";
				}
				if (string.IsNullOrEmpty(ChkToDt))
				{
					ChkToDt = "21000101";
				}

				WinTitle = "[PS_GA167] 레포트";
				ReportName = "PS_GA167_05.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@Ctgr", Ctgr));
				dataPackParameter.Add(new PSH_DataPackClass("@ModelNm", ModelNm));
				dataPackParameter.Add(new PSH_DataPackClass("@LocCls", LocCls));
				dataPackParameter.Add(new PSH_DataPackClass("@MngNo", MngNo));
				dataPackParameter.Add(new PSH_DataPackClass("@ChkFrDt", DateTime.ParseExact(ChkFrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ChkToDt", DateTime.ParseExact(ChkToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ChkCls", ChkCls));

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

					if (pVal.ItemUID == "BtnSrch01")
					{

						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_GA167_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_GA167_MTX02();
						}
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_GA167_MTX03();
						}
					}
					else if (pVal.ItemUID == "BtnSrch05")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_GA167_MTX05();
						}
					}
					else if (pVal.ItemUID == "BtnPrt01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_GA167_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "BtnPrt05")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_GA167_Print_Report05);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
					}
					if (pVal.ItemUID == "Folder04")
					{
						oForm.PaneLevel = 4;
						oForm.DefButton = "BtnSrch04";
					}
					if (pVal.ItemUID == "Folder05")
					{
						oForm.PaneLevel = 5;
						oForm.DefButton = "BtnSrch05";
					}
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "TeamCode01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MngNo01", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode02", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MngNo02", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode03", "");

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode05", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MngNo05", "");
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
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_GA167_FlushToItemValue(pVal.ItemUID);
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
					PS_GA167_FlushToItemValue(pVal.ItemUID);
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
					PS_GA167_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid05);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA167A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA167B);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA167C);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA167D);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA167E);
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
