using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
    /// 제안평가(QC)
    /// </summary>
	internal class PS_QM152 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM152.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM152_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM152");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_QM152_CreateItems();
				PS_QM152_ComboBox_Setting();

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
				oForm.ActiveItem = "ym"; //최초 커서위치
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_QM152_CreateItems
		/// </summary>
		private void PS_QM152_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				//접수일자
				oForm.DataSources.UserDataSources.Add("ymd", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ymd").Specific.DataBind.SetBound(true, "", "ymd");
				oForm.DataSources.UserDataSources.Item("ymd").Value = DateTime.Now.ToString("yyyyMMdd");

				//사원번호
				oForm.DataSources.UserDataSources.Add("sabun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("sabun").Specific.DataBind.SetBound(true, "", "sabun");
				oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

				//제안번호
				oForm.DataSources.UserDataSources.Add("proposalno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("proposalno").Specific.DataBind.SetBound(true, "", "proposalno");

				//소속
				oForm.DataSources.UserDataSources.Add("buseonm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("buseonm").Specific.DataBind.SetBound(true, "", "buseonm");
				oForm.DataSources.UserDataSources.Add("sectnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("sectnm").Specific.DataBind.SetBound(true, "", "sectnm");
				oForm.DataSources.UserDataSources.Add("staffnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("staffnm").Specific.DataBind.SetBound(true, "", "staffnm");
				oForm.DataSources.UserDataSources.Add("wrkgrdnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("wrkgrdnm").Specific.DataBind.SetBound(true, "", "wrkgrdnm");

				//제목
				oForm.DataSources.UserDataSources.Add("title", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("title").Specific.DataBind.SetBound(true, "", "title");

				//실시부서,담당
				oForm.DataSources.UserDataSources.Add("sbuseo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("sbuseo").Specific.DataBind.SetBound(true, "", "sbuseo");
				oForm.DataSources.UserDataSources.Add("ssect", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ssect").Specific.DataBind.SetBound(true, "", "ssect");

				oForm.DataSources.UserDataSources.Add("sbuseonm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("sbuseonm").Specific.DataBind.SetBound(true, "", "sbuseonm");
				oForm.DataSources.UserDataSources.Add("ssectnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ssectnm").Specific.DataBind.SetBound(true, "", "ssectnm");

				//시행일자
				oForm.DataSources.UserDataSources.Add("isymd", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("isymd").Specific.DataBind.SetBound(true, "", "isymd");

				//효과금액
				oForm.DataSources.UserDataSources.Add("effectamt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("effectamt").Specific.DataBind.SetBound(true, "", "effectamt");

				//점수
				oForm.DataSources.UserDataSources.Add("mark", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("mark").Specific.DataBind.SetBound(true, "", "mark");

				//등급
				oForm.DataSources.UserDataSources.Add("grade", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("grade").Specific.DataBind.SetBound(true, "", "grade");

				//평가점
				oForm.DataSources.UserDataSources.Add("par", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("par").Specific.DataBind.SetBound(true, "", "par");

				//시상액
				oForm.DataSources.UserDataSources.Add("prizeamt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("prizeamt").Specific.DataBind.SetBound(true, "", "prizeamt");

				oForm.DataSources.UserDataSources.Add("mark_1", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("mark_1").Specific.DataBind.SetBound(true, "", "mark_1");

				oForm.DataSources.UserDataSources.Add("grade_1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("grade_1").Specific.DataBind.SetBound(true, "", "grade_1");

				oForm.DataSources.UserDataSources.Add("par_1", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("par_1").Specific.DataBind.SetBound(true, "", "par_1");

				oForm.DataSources.UserDataSources.Add("prizeamt_1", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("prizeamt_1").Specific.DataBind.SetBound(true, "", "prizeamt_1");

				oForm.DataSources.UserDataSources.Add("mark_a", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("mark_a").Specific.DataBind.SetBound(true, "", "mark_a");

				oForm.DataSources.UserDataSources.Add("grade_a", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("grade_a").Specific.DataBind.SetBound(true, "", "grade_a");

				oForm.DataSources.UserDataSources.Add("par_a", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("par_a").Specific.DataBind.SetBound(true, "", "par_a");

				oForm.DataSources.UserDataSources.Add("prizeamt_a", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("prizeamt_a").Specific.DataBind.SetBound(true, "", "prizeamt_a");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM152_ComboBox_Setting
		/// </summary>
		private void PS_QM152_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장콤보박스세팅
				dataHelpClass.Set_ComboList(oForm.Items.Item("saup").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("saup").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//제안구분
				oForm.Items.Item("div").Specific.ValidValues.Add("0", "정식");
				oForm.Items.Item("div").Specific.ValidValues.Add("1", "약식");
				oForm.Items.Item("div").Specific.ValidValues.Add("2", "등외");
				oForm.Items.Item("div").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//분야
				dataHelpClass.Set_ComboList(oForm.Items.Item("field").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q015' order by U_Minor", "", false, false);
				oForm.Items.Item("field").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//채택구분
				oForm.Items.Item("adoptdiv").Specific.ValidValues.Add("Y", "채택");
				oForm.Items.Item("adoptdiv").Specific.ValidValues.Add("N", "불채택");
				oForm.Items.Item("adoptdiv").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//실시구분
				oForm.Items.Item("isdiv").Specific.ValidValues.Add("N", "미실시");
				oForm.Items.Item("isdiv").Specific.ValidValues.Add("Y", "실시");
				oForm.Items.Item("isdiv").Specific.ValidValues.Add("G", "검토");
				oForm.Items.Item("isdiv").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 그리드에 데이터 로드 (조회)
		/// </summary>
		private void PS_QM152_MTX01()
		{
			string saup;
			string YM;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim();
				YM = oForm.Items.Item("ym").Specific.Value.ToString().Trim();

				sQry = " EXEC [PS_QM152_01]  '" + saup + "', '" + YM + "'";

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				oRecordSet.DoQuery(sQry);
				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다. 사업장,년월을 확인 하세요.";
					throw new Exception();
				}

				oGrid.AutoResizeColumns();
				oForm.Update();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 그리드 자료를 head에 로드
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM152_MTX02(string oUID, int oRow, string oCol)
		{
			int sRow;
			string Param01;
			string Param02;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				sRow = oRow;
				Param01 = oGrid.DataTable.Columns.Item("사업장").Cells.Item(oRow).Value.ToString().Trim();
				Param02 = oGrid.DataTable.Columns.Item("제안번호").Cells.Item(oRow).Value.ToString().Trim();

				sQry = "EXEC PS_QM152_02 '" + Param01 + "', '" + Param02 + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					PS_QM152_Form_ini();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oForm.DataSources.UserDataSources.Item("ymd").Value = Convert.ToDateTime(oRecordSet.Fields.Item("ymd").Value.ToString().Trim()).ToString("yyyyMMdd");
				oForm.DataSources.UserDataSources.Item("sabun").Value = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("proposalno").Value = oRecordSet.Fields.Item("proposalno").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("buseonm").Value = oRecordSet.Fields.Item("buseonm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("sectnm").Value = oRecordSet.Fields.Item("sectnm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("staffnm").Value = oRecordSet.Fields.Item("staffnm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("wrkgrdnm").Value = oRecordSet.Fields.Item("wrkgrdnm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("title").Value = oRecordSet.Fields.Item("title").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("sbuseo").Value = oRecordSet.Fields.Item("sbuseo").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("ssect").Value = oRecordSet.Fields.Item("ssect").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("sbuseonm").Value = oRecordSet.Fields.Item("sbuseonm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("ssectnm").Value = oRecordSet.Fields.Item("ssectnm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("isymd").Value = Convert.ToDateTime(oRecordSet.Fields.Item("isymd").Value.ToString().Trim()).ToString("yyyyMMdd");
				oForm.DataSources.UserDataSources.Item("effectamt").Value = oRecordSet.Fields.Item("effectamt").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("mark").Value = oRecordSet.Fields.Item("mark").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("grade").Value = oRecordSet.Fields.Item("grade").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("par").Value = oRecordSet.Fields.Item("par").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("prizeamt").Value = oRecordSet.Fields.Item("prizeamt").Value.ToString().Trim();

				oForm.DataSources.UserDataSources.Item("mark_1").Value = oRecordSet.Fields.Item("mark_1").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("grade_1").Value = oRecordSet.Fields.Item("grade_1").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("par_1").Value = oRecordSet.Fields.Item("par_1").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("prizeamt_1").Value = oRecordSet.Fields.Item("prizeamt_1").Value.ToString().Trim();

				oForm.DataSources.UserDataSources.Item("mark_a").Value = oRecordSet.Fields.Item("mark_a").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("grade_a").Value = oRecordSet.Fields.Item("grade_a").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("par_a").Value = oRecordSet.Fields.Item("par_a").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("prizeamt_a").Value = oRecordSet.Fields.Item("prizeamt_a").Value.ToString().Trim();

				oForm.Items.Item("saup").Specific.Select(oRecordSet.Fields.Item("saup").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("div").Specific.Select(oRecordSet.Fields.Item("div").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("field").Specific.Select(oRecordSet.Fields.Item("field").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("adoptdiv").Specific.Select(oRecordSet.Fields.Item("adoptdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("isdiv").Specific.Select(oRecordSet.Fields.Item("isdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.ActiveItem = "adoptdiv";
				oForm.Items.Item("proposalno").Enabled = false;

				oForm.Update();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 데이타 저장
		/// </summary>
		private void PS_QM152_SAVE()
		{
			string grade_1;
			string adoptdiv;
			string saup;
			string proposalno;
			string isdiv;
			string grade_a;
			decimal mark_a;
			decimal Par_1;
			decimal mark_1;
			decimal PrizeAmt_1;
			decimal Par_a;
			decimal PrizeAmt_a;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim(); ;
				proposalno = oForm.Items.Item("proposalno").Specific.Value.ToString().Trim();
				adoptdiv = oForm.Items.Item("adoptdiv").Specific.Value.ToString().Trim();
				isdiv = oForm.Items.Item("isdiv").Specific.Value.ToString().Trim();
				grade_1 = oForm.Items.Item("grade_1").Specific.Value.ToString().Trim();
				grade_a = oForm.Items.Item("grade_a").Specific.Value.ToString().Trim();

				mark_1 = Convert.ToDecimal(oForm.Items.Item("mark_1").Specific.Value.ToString().Trim());
				Par_1 = Convert.ToDecimal(oForm.Items.Item("par_1").Specific.Value.ToString().Trim());
				PrizeAmt_1 = Convert.ToDecimal(oForm.Items.Item("prizeamt_1").Specific.Value.ToString().Trim());
				mark_a = Convert.ToDecimal(oForm.Items.Item("mark_a").Specific.Value.ToString().Trim());
				Par_a = Convert.ToDecimal(oForm.Items.Item("par_a").Specific.Value.ToString().Trim());
				PrizeAmt_a = Convert.ToDecimal(oForm.Items.Item("prizeamt_a").Specific.Value.ToString().Trim());

				//갱신
				sQry = " Update [ZPS_QM151] set ";
				sQry += "adoptdiv = '" + adoptdiv + "',";
				sQry += "isdiv = '" + isdiv + "',";
				sQry += "mark_1 = '" + mark_1 + "',";
				sQry += "grade_1 = '" + grade_1 + "',";
				sQry += "par_1 = '" + Par_1 + "',";
				sQry += "prizeamt_1 = '" + PrizeAmt_1 + "',";
				sQry += "mark_a = '" + mark_a + "',";
				sQry += "grade_a = '" + grade_a + "',";
				sQry += "par_a = '" + Par_a + "',";
				sQry += "prizeamt_a = '" + PrizeAmt_a + "'";
				sQry += " Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
				oRecordSet.DoQuery(sQry);

				PS_QM152_MTX01();
				PS_QM152_Form_ini();
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
		/// PS_QM152_Form_ini
		/// </summary>
		private void PS_QM152_Form_ini()
		{
			try
			{
				oForm.Freeze(true);
				oForm.DataSources.UserDataSources.Item("ymd").Value = "";
				oForm.DataSources.UserDataSources.Item("sabun").Value = "";
				oForm.DataSources.UserDataSources.Item("kname").Value = "";
				oForm.DataSources.UserDataSources.Item("proposalno").Value = "";
				oForm.DataSources.UserDataSources.Item("buseonm").Value = "";
				oForm.DataSources.UserDataSources.Item("sectnm").Value = "";
				oForm.DataSources.UserDataSources.Item("staffnm").Value = "";
				oForm.DataSources.UserDataSources.Item("wrkgrdnm").Value = "";
				oForm.DataSources.UserDataSources.Item("sbuseo").Value = "";
				oForm.DataSources.UserDataSources.Item("ssect").Value = "";
				oForm.DataSources.UserDataSources.Item("sbuseonm").Value = "";
				oForm.DataSources.UserDataSources.Item("ssectnm").Value = "";
				oForm.DataSources.UserDataSources.Item("isymd").Value = "";
				oForm.DataSources.UserDataSources.Item("effectamt").Value = "0";
				oForm.DataSources.UserDataSources.Item("mark").Value = "0";
				oForm.DataSources.UserDataSources.Item("grade").Value = "";
				oForm.DataSources.UserDataSources.Item("par").Value = "0";
				oForm.DataSources.UserDataSources.Item("prizeamt").Value = "0";
				oForm.DataSources.UserDataSources.Item("mark_1").Value = "0";
				oForm.DataSources.UserDataSources.Item("grade_1").Value = "";
				oForm.DataSources.UserDataSources.Item("par_1").Value = "0";
				oForm.DataSources.UserDataSources.Item("prizeamt_1").Value = "0";
				oForm.DataSources.UserDataSources.Item("mark_a").Value = "0";
				oForm.DataSources.UserDataSources.Item("grade_a").Value = "";
				oForm.DataSources.UserDataSources.Item("par_a").Value = "0";
				oForm.DataSources.UserDataSources.Item("prizeamt_a").Value = "0";

				oForm.Items.Item("div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("field").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("title").Specific.Value = "";
				oForm.Items.Item("adoptdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("isdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.ActiveItem = "proposalno";
				oForm.Update();
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
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
					if (pVal.ItemUID == "Btn_ret")
					{
						PS_QM152_MTX01();
					}
					if (pVal.ItemUID == "Btn_save")
					{
						PS_QM152_SAVE();
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "proposalno")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("proposalno").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
		}

		/// <summary>
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Div;
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Div")
					{
						Div = oForm.Items.Item("Div").Specific.Value.ToString().Trim();

						if (Div == "0")
						{
							oForm.Items.Item("DivNm").Specific.Value = "정식";
						}
						else if (Div == "1")
						{
							oForm.Items.Item("DivNm").Specific.Value = "약식";
						}
						else if (Div == "2")
						{
							oForm.Items.Item("DivNm").Specific.Value = "등외";
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
					if (pVal.ItemUID == "Grid01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.Row >= 0)
							{
								PS_QM152_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
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
			double Mark;
			double mark_1;
			double mark_c;
			string saup;
			string proposalno;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "proposalno") //제안번호
						{
							saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim();
							proposalno = oForm.Items.Item("proposalno").Specific.Value.ToString().Trim();
							sQry = "EXEC PS_QM152_02 '" + saup + "', '" + proposalno + "'";
							oRecordSet.DoQuery(sQry);

							oForm.DataSources.UserDataSources.Item("ymd").Value = Convert.ToDateTime(oRecordSet.Fields.Item("ymd").Value.ToString().Trim()).ToString("yyyyMMdd");
							oForm.DataSources.UserDataSources.Item("sabun").Value = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("proposalno").Value = oRecordSet.Fields.Item("proposalno").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("buseonm").Value = oRecordSet.Fields.Item("buseonm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("sectnm").Value = oRecordSet.Fields.Item("sectnm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("staffnm").Value = oRecordSet.Fields.Item("staffnm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("wrkgrdnm").Value = oRecordSet.Fields.Item("wrkgrdnm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("title").Value = oRecordSet.Fields.Item("title").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("sbuseo").Value = oRecordSet.Fields.Item("sbuseo").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("ssect").Value = oRecordSet.Fields.Item("ssect").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("sbuseonm").Value = oRecordSet.Fields.Item("sbuseonm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("ssectnm").Value = oRecordSet.Fields.Item("ssectnm").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("isymd").Value = Convert.ToDateTime(oRecordSet.Fields.Item("isymd").Value.ToString().Trim()).ToString("yyyyMMdd");
							oForm.DataSources.UserDataSources.Item("effectamt").Value = oRecordSet.Fields.Item("effectamt").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("mark").Value = oRecordSet.Fields.Item("mark").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("grade").Value = oRecordSet.Fields.Item("grade").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("par").Value = oRecordSet.Fields.Item("par").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("prizeamt").Value = oRecordSet.Fields.Item("prizeamt").Value.ToString().Trim();

							oForm.DataSources.UserDataSources.Item("mark_1").Value = oRecordSet.Fields.Item("mark_1").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("grade_1").Value = oRecordSet.Fields.Item("grade_1").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("par_1").Value = oRecordSet.Fields.Item("par_1").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("prizeamt_1").Value = oRecordSet.Fields.Item("prizeamt_1").Value.ToString().Trim();

							oForm.DataSources.UserDataSources.Item("mark_a").Value = oRecordSet.Fields.Item("mark_a").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("grade_a").Value = oRecordSet.Fields.Item("grade_a").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("par_a").Value = oRecordSet.Fields.Item("par_a").Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("prizeamt_a").Value = oRecordSet.Fields.Item("prizeamt_a").Value.ToString().Trim();

							oForm.Items.Item("saup").Specific.Select(oRecordSet.Fields.Item("saup").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("div").Specific.Select(oRecordSet.Fields.Item("div").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("field").Specific.Select(oRecordSet.Fields.Item("field").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

							oForm.Items.Item("adoptdiv").Specific.Select(oRecordSet.Fields.Item("adoptdiv").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("isdiv").Specific.Select(oRecordSet.Fields.Item("isdiv").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

							oForm.ActiveItem = "adoptdiv"; //커서위치
						}
						
						if (pVal.ItemUID == "mark_1") //점수QC
						{
							sQry = "SELECT Grade, Par, PrizeAmt FROM [ZPS_QM150] WHERE '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "' BETWEEN MarkMin AND MarkMax AND BPLID = '" + oForm.Items.Item("saup").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("grade_1").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("par_1").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							oForm.Items.Item("prizeamt_1").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();

							Mark = Convert.ToDouble(oForm.Items.Item("mark").Specific.Value.ToString().Trim());
							mark_1 = Convert.ToDouble(oForm.Items.Item("mark_1").Specific.Value.ToString().Trim());
							mark_c = Convert.ToDouble(Math.Round((Mark + mark_1) / 2, 0));
							oForm.Items.Item("mark_a").Specific.Value = Convert.ToString(mark_c);

							sQry = "SELECT Grade, Par, PrizeAmt FROM [ZPS_QM150] WHERE '" + mark_c + "' BETWEEN MarkMin AND MarkMax AND BPLID = '" + oForm.Items.Item("saup").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("grade_a").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("par_a").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							oForm.Items.Item("prizeamt_a").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
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
							PS_QM152_Form_ini();
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
