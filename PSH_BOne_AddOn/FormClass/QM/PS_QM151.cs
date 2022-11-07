using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 제안등록
	/// </summary>
	internal class PS_QM151 : PSH_BaseClass
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM151.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM151_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM151");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM151_CreateItems();
				PS_QM151_ComboBox_Setting();

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
				oForm.ActiveItem = "ymd"; //최초 커서위치
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_QM151_CreateItems
		/// </summary>
		private void PS_QM151_CreateItems()
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
				oForm.DataSources.UserDataSources.Add("buseo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("buseo").Specific.DataBind.SetBound(true, "", "buseo");
				oForm.DataSources.UserDataSources.Add("sect", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("sect").Specific.DataBind.SetBound(true, "", "sect");
				oForm.DataSources.UserDataSources.Add("staff", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("staff").Specific.DataBind.SetBound(true, "", "staff");
				oForm.DataSources.UserDataSources.Add("wrkgrd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("wrkgrd").Specific.DataBind.SetBound(true, "", "wrkgrd");

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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM151_ComboBox_Setting
		/// </summary>
		private void PS_QM151_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("saup").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("saup").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//제안구분
				oForm.Items.Item("div").Specific.ValidValues.Add("", "선택");
				oForm.Items.Item("div").Specific.ValidValues.Add("0", "정식");
				oForm.Items.Item("div").Specific.ValidValues.Add("1", "약식");
				oForm.Items.Item("div").Specific.ValidValues.Add("2", "등외");
				oForm.Items.Item("div").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//분야
				oForm.Items.Item("field").Specific.ValidValues.Add("", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("field").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q015' order by U_Minor", "", false, false);
				oForm.Items.Item("field").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//채택구분
				oForm.Items.Item("adoptdiv").Specific.ValidValues.Add("Y", "채택");
				oForm.Items.Item("adoptdiv").Specific.ValidValues.Add("N", "불채택");
				oForm.Items.Item("adoptdiv").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//실시구분
				oForm.Items.Item("isdiv").Specific.ValidValues.Add("", "선택");
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
		private void PS_QM151_MTX01()
		{
			string saup;
			string YM;
			string sabun;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim();
				YM = oForm.Items.Item("ymd").Specific.Value.ToString().Trim().Substring(0, 6);
				sabun = oForm.Items.Item("sabun").Specific.Value.ToString().Trim();

				sQry = " EXEC [PS_QM151_01]  '" + saup + "', '" + YM + "', '" + sabun + "'";

				oGrid.DataTable.Clear();

				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				oRecordSet.DoQuery(sQry);
				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다. 사업장,접수일자,사원번호를 확인 하세요.";
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
		private void PS_QM151_MTX02(string oUID, int oRow, string oCol)
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

				sQry = "EXEC PS_QM151_02 '" + Param01 + "', '" + Param02 + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					PS_QM151_Form_ini();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oForm.DataSources.UserDataSources.Item("ymd").Value = Convert.ToDateTime(oRecordSet.Fields.Item("ymd").Value.ToString().Trim()).ToString("yyyyMMdd");
				oForm.DataSources.UserDataSources.Item("sabun").Value = oRecordSet.Fields.Item("sabun").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("proposalno").Value = oRecordSet.Fields.Item("proposalno").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("buseo").Value = oRecordSet.Fields.Item("buseo").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("sect").Value = oRecordSet.Fields.Item("sect").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("staff").Value = oRecordSet.Fields.Item("staff").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("wrkgrd").Value = oRecordSet.Fields.Item("wrkgrd").Value.ToString().Trim();
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

				oForm.Items.Item("saup").Specific.Select(oRecordSet.Fields.Item("saup").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("div").Specific.Select(oRecordSet.Fields.Item("div").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("field").Specific.Select(oRecordSet.Fields.Item("field").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("adoptdiv").Specific.Select(oRecordSet.Fields.Item("adoptdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("isdiv").Specific.Select(oRecordSet.Fields.Item("isdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.ActiveItem = "ymd";
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
		private void PS_QM151_SAVE()
		{
			int Seqncom;
			string YM;
			string staffnm;
			string buseonm;
			string staff;
			string buseo;
			string kname;
			string ymd;
			string saup;
			string sabun;
			string proposalno;
			string Sect;
			string wrkgrd;
			string sectnm;
			string wrkgrdnm;
			string isymd;
			string ssect;
			string sbuseo;
			string adoptdiv;
			string field;
			string Div;
			string Title;
			string isdiv;
			string sbuseonm;
			string ssectnm;
			string Grade;
			object Mark;
			object effectamt;
			double Par;
			double PrizeAmt;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim();
				ymd = oForm.Items.Item("ymd").Specific.Value.ToString().Trim();
				sabun = oForm.Items.Item("sabun").Specific.Value.ToString().Trim();
				proposalno = oForm.Items.Item("proposalno").Specific.Value.ToString().Trim();
				Div = oForm.Items.Item("div").Specific.Value.ToString().Trim();
				field = oForm.Items.Item("field").Specific.Value.ToString().Trim();
				Title = oForm.Items.Item("title").Specific.Value.ToString().Trim();
				adoptdiv = oForm.Items.Item("adoptdiv").Specific.Value.ToString().Trim();
				isdiv = oForm.Items.Item("isdiv").Specific.Value.ToString().Trim();
				sbuseo = oForm.Items.Item("sbuseo").Specific.Value.ToString().Trim();
				ssect = oForm.Items.Item("ssect").Specific.Value.ToString().Trim();
				sbuseonm = oForm.Items.Item("sbuseonm").Specific.Value.ToString().Trim();
				ssectnm = oForm.Items.Item("ssectnm").Specific.Value.ToString().Trim();
				isymd = oForm.Items.Item("isymd").Specific.Value.ToString().Trim();
				effectamt = oForm.Items.Item("effectamt").Specific.Value.ToString().Trim();
				Mark = oForm.Items.Item("mark").Specific.Value.ToString().Trim();
				Grade = oForm.Items.Item("grade").Specific.Value.ToString().Trim();
				Par = Convert.ToDouble(oForm.Items.Item("par").Specific.Value.ToString().Trim());
				PrizeAmt = Convert.ToDouble(oForm.Items.Item("prizeamt").Specific.Value.ToString().Trim());

				sQry = " SELECT U_FullName,  ";
				sQry += "U_TeamCode, TeamName = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '1' And U_Code = U_TeamCode),''), ";
				sQry += "U_RspCode,  RspName  = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '2' And U_Code = U_RspCode),''), ";
				sQry += "U_ClsCode,  ClsName  = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '9' And U_Code  = U_ClsCode),''), ";
				sQry += "U_JigCod,   JigName  = (SELECT U_CodeNm From [@PS_HR200L] WHERE Code = 'P129' And U_Code = U_JigCod) ";
				sQry += "FROM [@PH_PY001A] WHERE Code = '" + sabun + "'";
				oRecordSet.DoQuery(sQry);

				kname = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				buseo = oRecordSet.Fields.Item(1).Value.ToString().Trim();
				buseonm = oRecordSet.Fields.Item(2).Value.ToString().Trim();
				Sect = oRecordSet.Fields.Item(3).Value.ToString().Trim();
				sectnm = oRecordSet.Fields.Item(4).Value.ToString().Trim();
				staff = oRecordSet.Fields.Item(5).Value.ToString().Trim();
				staffnm = oRecordSet.Fields.Item(6).Value.ToString().Trim();
				wrkgrd = oRecordSet.Fields.Item(7).Value.ToString().Trim();
				wrkgrdnm = oRecordSet.Fields.Item(8).Value.ToString().Trim();

				if (string.IsNullOrEmpty(saup))
				{
					errMessage = "사업장에러. 확인바랍니다..";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(ymd))
				{
					errMessage = "접수일자에러. 확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(sabun))
				{
					errMessage = "사원번호에러. 확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(Div))
				{
					errMessage = "제안구분 선택하세요. 확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(Title))
				{
					errMessage = "제안분야 선택하세요. 확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(isdiv))
				{
					errMessage = "실시구분 선택하세요. 확인바랍니다.";
					throw new Exception();
				}

				sQry = " Select Count(*) From [ZPS_QM151] Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
				{
					//갱신
					sQry = "Update [ZPS_QM151] set ";
					sQry += "ymd = '" + ymd + "',";
					sQry += "sabun = '" + sabun + "',";
					sQry += "kname = '" + kname + "',";
					sQry += "buseo = '" + buseo + "',";
					sQry += "sect = '" + Sect + "',";
					sQry += "staff = '" + staff + "',";
					sQry += "wrkgrd = '" + wrkgrd + "',";
					sQry += "buseonm = '" + buseonm + "',";
					sQry += "sectnm = '" + sectnm + "',";
					sQry += "staffnm = '" + staffnm + "',";
					sQry += "wrkgrdnm = '" + wrkgrdnm + "',";
					sQry += "div = '" + Div + "',";
					sQry += "field = '" + field + "',";
					sQry += "title = '" + Title + "',";
					sQry += "adoptdiv = '" + adoptdiv + "',";
					sQry += "isdiv = '" + isdiv + "',";
					sQry += "sbuseo = '" + sbuseo + "',";
					sQry += "ssect = '" + ssect + "',";
					sQry += "sbuseonm = '" + sbuseonm + "',";
					sQry += "ssectnm = '" + ssectnm + "',";
					sQry += "isymd = '" + isymd + "',";
					sQry += "effectamt = '" + effectamt + "',";
					sQry += "mark = '" + Mark + "',";
					sQry += "grade = '" + Grade + "',";
					sQry += "par = '" + Par + "',";
					sQry += "prizeamt = '" + PrizeAmt + "'";
					sQry += " Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
					oRecordSet.DoQuery(sQry);
				}
				else
				{
					//순번 계산
					YM = ymd.Substring(0, 6);
					sQry = " Select Convert(int,Right(Isnull(Max(proposalno),''),3)) From [ZPS_QM151] Where saup = '" + saup + "' And Convert(char(6),ymd,112) Like '" + YM + "' + '%'";
					oRecordSet.DoQuery(sQry);

					Seqncom = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());
					Seqncom += 1;
					proposalno = ymd.Substring(0, 6) + Seqncom.ToString().PadLeft(3, '0');
					oForm.Items.Item("proposalno").Specific.Value = proposalno;

					sQry = "INSERT INTO [ZPS_QM151]";
					sQry += " (";
					sQry += "saup,";
					sQry += "proposalno,";
					sQry += "ymd,";
					sQry += "sabun,";
					sQry += "kname,";
					sQry += "buseo,";
					sQry += "sect,";
					sQry += "staff,";
					sQry += "wrkgrd,";
					sQry += "buseonm,";
					sQry += "sectnm,";
					sQry += "staffnm,";
					sQry += "wrkgrdnm,";
					sQry += "div,";
					sQry += "field,";
					sQry += "title,";
					sQry += "adoptdiv,";
					sQry += "isdiv,";
					sQry += "sbuseo,";
					sQry += "ssect,";
					sQry += "sbuseonm,";
					sQry += "ssectnm,";
					sQry += "isymd,";
					sQry += "effectamt,";
					sQry += "mark,";
					sQry += "grade,";
					sQry += "par,";
					sQry += "prizeamt";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += "'" + saup + "',";
					sQry += "'" + proposalno + "',";
					sQry += "'" + ymd + "',";
					sQry += "'" + sabun + "',";
					sQry += "'" + kname + "',";
					sQry += "'" + buseo + "',";
					sQry += "'" + Sect + "',";
					sQry += "'" + staff + "',";
					sQry += "'" + wrkgrd + "',";
					sQry += "'" + buseonm + "',";
					sQry += "'" + sectnm + "',";
					sQry += "'" + staffnm + "',";
					sQry += "'" + wrkgrdnm + "',";
					sQry += "'" + Div + "',";
					sQry += "'" + field + "',";
					sQry += "'" + Title + "',";
					sQry += "'" + adoptdiv + "',";
					sQry += "'" + isdiv + "',";
					sQry += "'" + sbuseo + "',";
					sQry += "'" + ssect + "',";
					sQry += "'" + sbuseonm + "',";
					sQry += "'" + ssectnm + "',";
					sQry += "'" + isymd + "',";
					sQry += effectamt + ",";
					sQry += Mark + ",";
					sQry += "'" + Grade + "',";
					sQry += Par + ",";
					sQry += PrizeAmt + "";
					sQry += " ) ";
					oRecordSet.DoQuery(sQry);
				}
				PS_QM151_MTX01();
				PS_QM151_Form_ini();
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
		/// 선택된 자료 삭제
		/// </summary>
		private void PS_QM151_Delete()
		{
			string saup;
			string proposalno;
			int Cnt;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				saup = oForm.Items.Item("saup").Specific.Value.ToString().Trim();
				proposalno = oForm.Items.Item("proposalno").Specific.Value.ToString().Trim();

				sQry = " Select Count(*) From [ZPS_QM151] Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
				oRecordSet.DoQuery(sQry);
				Cnt = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());

				if (Cnt > 0)
				{
					if (PSH_Globals.SBO_Application.MessageBox(" 선택한라인을 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
					{
						sQry = " Select isnull(mark_1,0) From [ZPS_QM151] Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
						oRecordSet.DoQuery(sQry);

						if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
						{
							sQry = "Delete From [ZPS_QM151] Where saup = '" + saup + "' And proposalno = '" + proposalno + "'";
							oRecordSet.DoQuery(sQry);
							PS_QM151_MTX01();
							PS_QM151_Form_ini();
						}
						else
						{
							errMessage = "QC점수가 부여된 자료 입니다. 확인하세요.";
							throw new Exception();
						}
					}
				}
				else
				{
					errMessage = "조회후 삭제하십시요.";
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 화면청소
		/// </summary>
		private void PS_QM151_Form_ini()
		{
			try
			{
				oForm.DataSources.UserDataSources.Item("ymd").Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.DataSources.UserDataSources.Item("sabun").Value = "";
				oForm.DataSources.UserDataSources.Item("kname").Value = "";
				oForm.DataSources.UserDataSources.Item("proposalno").Value = "";
				oForm.DataSources.UserDataSources.Item("buseo").Value = "";
				oForm.DataSources.UserDataSources.Item("sect").Value = "";
				oForm.DataSources.UserDataSources.Item("staff").Value = "";
				oForm.DataSources.UserDataSources.Item("wrkgrd").Value = "";
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
				oForm.Items.Item("div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("field").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("title").Specific.Value = "";
				oForm.Items.Item("adoptdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("isdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.ActiveItem = "ymd";
				oForm.Update();
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
						PS_QM151_MTX01();
					}
					else if (pVal.ItemUID == "Btn_save")
					{
						PS_QM151_SAVE();
					}
					else if (pVal.ItemUID == "Btn_delete")
					{
						PS_QM151_Delete();
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
						if (pVal.ItemUID == "sbuseo")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("sbuseo").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ssect")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ssect").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "sabun")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("sabun").Specific.Value.ToString().Trim()))
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
								PS_QM151_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
						if (pVal.ItemUID == "sabun") //사원명
						{
							sQry = " SELECT U_FullName,  ";
							sQry += "U_TeamCode, TeamName = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '1' And U_Code = U_TeamCode),''), ";
							sQry += "U_RspCode,  RspName  = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '2' And U_Code = U_RspCode),''), ";
							sQry += "U_ClsCode,  ClsName  = Isnull((SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '9' And U_Code  = U_ClsCode),''), ";
							sQry += "U_JigCod,   JigName  = (SELECT U_CodeNm From [@PS_HR200L] WHERE Code = 'P129' And U_Code = U_JigCod) ";
							sQry += "FROM [@PH_PY001A] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("kname").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("buseo").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							oForm.Items.Item("buseonm").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
							oForm.Items.Item("sect").Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
							oForm.Items.Item("sectnm").Specific.Value = oRecordSet.Fields.Item(4).Value.ToString().Trim();
							oForm.Items.Item("staff").Specific.Value = oRecordSet.Fields.Item(5).Value.ToString().Trim();
							oForm.Items.Item("staffnm").Specific.Value = oRecordSet.Fields.Item(6).Value.ToString().Trim();
							oForm.Items.Item("wrkgrd").Specific.Value = oRecordSet.Fields.Item(7).Value.ToString().Trim();
							oForm.Items.Item("wrkgrdnm").Specific.Value = oRecordSet.Fields.Item(8).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "sbuseo") //실시부서
						{
							sQry = "SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '1' And U_Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("sbuseonm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "ssect") //실시담당
						{
							sQry = "SELECT U_CodeNm From [@PS_HR200L] WHERE Code = '2' And U_Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("ssectnm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "mark") //점수
						{
							sQry = "SELECT Grade, Par, PrizeAmt FROM [ZPS_QM150] WHERE '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "' BETWEEN MarkMin AND MarkMax AND BPLID = '" + oForm.Items.Item("saup").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("grade").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("par").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							oForm.Items.Item("prizeamt").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
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
							PS_QM151_Form_ini();
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
