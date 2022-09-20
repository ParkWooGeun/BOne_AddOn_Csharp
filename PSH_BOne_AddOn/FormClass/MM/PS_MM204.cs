using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검수현황
	/// </summary>
	internal class PS_MM204 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01 ; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM204.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_MM204_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM204");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);

                PS_MM204_CreateItems();
                PS_MM204_SetComboBox();

                //oForm.Items.Item("StrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// CreateItems
        /// </summary>
        private void PS_MM204_CreateItems()
        {
            try
            {
                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //기간(FR)
                oForm.DataSources.UserDataSources.Add("StrDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("StrDate").Specific.DataBind.SetBound(true, "", "StrDate");
                oForm.DataSources.UserDataSources.Item("StrDate").Value = DateTime.Now.ToString("yyyyMMdd");

                //기간(TO)
                oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");
                oForm.DataSources.UserDataSources.Item("EndDate").Value = DateTime.Now.ToString("yyyyMMdd");

                //품목그룹
                oForm.DataSources.UserDataSources.Add("ItmsGrpCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmsGrpCod").Specific.DataBind.SetBound(true, "", "ItmsGrpCod");

                //대분류(코드)
                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                //대분류(명)
                oForm.DataSources.UserDataSources.Add("BsortName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("BsortName").Specific.DataBind.SetBound(true, "", "BsortName");

                //중분류(코드)
                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                //중분류(명)
                oForm.DataSources.UserDataSources.Add("MsortName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("MsortName").Specific.DataBind.SetBound(true, "", "MsortName");

                //품목코드
                oForm.DataSources.UserDataSources.Add("SItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SItemCode").Specific.DataBind.SetBound(true, "", "SItemCode");

                //공정
                oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

                //거래처(코드)
                oForm.DataSources.UserDataSources.Add("SCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SCardCode").Specific.DataBind.SetBound(true, "", "SCardCode");

                //품목/서비스 콤보
                oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("DocType").Specific.DataBind.SetBound(true, "", "DocType");

                //AP송장여부
                oForm.DataSources.UserDataSources.Add("AP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("AP").Specific.DataBind.SetBound(true, "", "AP");

                //품의구분 콤보
                oForm.DataSources.UserDataSources.Add("OrdTyp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("OrdTyp").Specific.DataBind.SetBound(true, "", "OrdTyp");

                //품의형태 콤보
                oForm.DataSources.UserDataSources.Add("POType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("POType").Specific.DataBind.SetBound(true, "", "POType");

                //출력구분
                oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Rad01").Specific.ValOn = "A";
                oForm.Items.Item("Rad01").Specific.ValOff = "0";
                oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad01").Specific.Selected = true;

                oForm.Items.Item("Rad02").Specific.ValOn = "B";
                oForm.Items.Item("Rad02").Specific.ValOff = "0";
                oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "OptionDS");

                oForm.Items.Item("Rad03").Specific.ValOn = "C";
                oForm.Items.Item("Rad03").Specific.ValOff = "0";
                oForm.Items.Item("Rad03").Specific.DataBind.SetBound(true, "", "OptionDS");

                oForm.Items.Item("Rad04").Specific.ValOn = "D";
                oForm.Items.Item("Rad04").Specific.ValOff = "0";
                oForm.Items.Item("Rad04").Specific.DataBind.SetBound(true, "", "OptionDS");

                oForm.Items.Item("Rad04").Specific.GroupWith("Rad01");
                oForm.Items.Item("Rad04").Specific.GroupWith("Rad02");
                oForm.Items.Item("Rad04").Specific.GroupWith("Rad03");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ComboBox 기본값설정
        /// </summary>
        private void PS_MM204_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //품목그룹
                oForm.Items.Item("ItmsGrpCod").Specific.ValidValues.Add("0", "전체");

                sQry = "SELECT ItmsGrpCod, ItmsGrpNam From [OITB]";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("ItmsGrpCod").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("ItmsGrpCod").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //사업장
                sQry = "SELECT U_Minor, U_CdName From [@PS_SY001L] WHERE Code = 'C105' AND U_UseYN Like 'Y' ORDER BY U_Seq";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                ////공정
                //PS_MM204_SetCpCode(); //사업장 콤보박스 이벤트 발생되므로 COMBO_SELECT 에서 호출

                //AP송장 여부
                oForm.Items.Item("AP").Specific.ValidValues.Add("A", "전체");
                oForm.Items.Item("AP").Specific.ValidValues.Add("Y", "A/P송장 발행건");
                oForm.Items.Item("AP").Specific.ValidValues.Add("N", "A/P송장 미발행건");
                oForm.Items.Item("AP").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품의구분(2012.03.36 송명규 추가)
                oForm.Items.Item("OrdTyp").Specific.ValidValues.Add("0", "전체");

                sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("OrdTyp").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("OrdTyp").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //구매방식
                oForm.Items.Item("POType").Specific.ValidValues.Add("%", "전체");

                sQry = "SELECT Code, Name From [@PSH_RETYPE]";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("POType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("POType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목/서비스(2014.0.05.02 송명규 추가)
                oForm.Items.Item("DocType").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("DocType").Specific.ValidValues.Add("I", "품목");
                oForm.Items.Item("DocType").Specific.ValidValues.Add("S", "서비스");
                oForm.Items.Item("DocType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
        private void PS_MM204_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                switch (oUID)
                {
                    case "ItmBsort":
                        sQry = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("BsortName").Specific.Value = oRecordSet.Fields.Item("Name").Value.ToString().Trim();
                        break;
                    case "ItmMsort":
                        sQry = "SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" + oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("MsortName").Specific.Value = oRecordSet.Fields.Item("U_CodeName").Value.ToString().Trim();
                        break;
                    case "StrDate":
                    case "EndDate":
                        PS_MM204_SetCpCode();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_MM204_PrintReport()
        {
            string WinTitle;
            string ReportName = string.Empty;
            string BPLID;
            string StrDate;
            string EndDate;
            string SItemCode;
            string EItemCode;
            string CpCode;
            string SCardCode;
            string ECardCode;
            string APYN;
            string ItmBsort;
            string ItmMsort;
            string ItmsGrpCod;
            string OrdTyp;
            string OptBtnValue;
            string POType;
            string DocType;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                BPLID = oForm.DataSources.UserDataSources.Item("BPLId").Value.ToString().Trim() == "0" ? "%" : oForm.DataSources.UserDataSources.Item("BPLId").Value.ToString().Trim(); //사업장
                StrDate = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("StrDate").Value.ToString().Trim()) ? "19000101" : oForm.DataSources.UserDataSources.Item("StrDate").Value.ToString().Trim().Replace(".", ""); //기간(FR)
                EndDate = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("EndDate").Value.ToString().Trim()) ? "21001231" : oForm.DataSources.UserDataSources.Item("EndDate").Value.ToString().Trim().Replace(".", ""); //기간(TO)
                ItmsGrpCod = oForm.DataSources.UserDataSources.Item("ItmsGrpCod").Value.ToString().Trim(); //품목그룹
                ItmBsort = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("ItmBsort").Value.ToString().Trim()) ? "%" : oForm.DataSources.UserDataSources.Item("ItmBsort").Value.ToString().Trim(); //대분류
                ItmMsort = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("ItmMsort").Value.ToString().Trim()) ? "%" : oForm.DataSources.UserDataSources.Item("ItmMsort").Value.ToString().Trim(); //중분류
                SItemCode = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("SItemCode").Value.ToString().Trim()) ? "1" : oForm.DataSources.UserDataSources.Item("SItemCode").Value.ToString().Trim(); //품목코드
                EItemCode = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("SItemCode").Value.ToString().Trim()) ? "ZZZZZZZZ" : oForm.DataSources.UserDataSources.Item("SItemCode").Value.ToString().Trim(); //품목코드
                CpCode = oForm.DataSources.UserDataSources.Item("CpCode").Value.ToString().Trim(); //공정
                SCardCode = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("SCardCode").Value.ToString().Trim()) ? "1" : oForm.DataSources.UserDataSources.Item("SCardCode").Value.ToString().Trim(); //거래처
                ECardCode = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("SCardCode").Value.ToString().Trim()) ? "ZZZZZZZZ" : oForm.DataSources.UserDataSources.Item("SCardCode").Value.ToString().Trim(); //거래처
                DocType = oForm.DataSources.UserDataSources.Item("DocType").Value.ToString().Trim(); //품목/서비스
                APYN = oForm.DataSources.UserDataSources.Item("AP").Value.ToString().Trim(); //AP송장 여부
                OrdTyp = oForm.DataSources.UserDataSources.Item("OrdTyp").Value.ToString().Trim(); //품의구분
                POType = oForm.DataSources.UserDataSources.Item("POType").Value.ToString().Trim(); //품의형태
                OptBtnValue = oForm.DataSources.UserDataSources.Item("OptionDS").Value.ToString().Trim(); //출력구분

                WinTitle = "[PS_MM204] 검수현황";
                
                if (OptBtnValue == "A") //전체
                {
                    ReportName = "PS_MM204_01.rpt";
                }
                else if (OptBtnValue == "B") //AP송장별
                {
                    ReportName = "PS_MM204_02.rpt";
                }
                else if (OptBtnValue == "C") //거래처별
                {
                    ReportName = "PS_MM204_03.rpt";
                }
                else if (OptBtnValue == "D") //거래처발송
                {
                    ReportName = "PS_MM204_04.rpt";
                }
                //프로시저 : PS_MM204_00

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                //Formula                
                dataPackFormula.Add(new PSH_DataPackClass("@StrDate", StrDate == "19000101" ? "All" : dataHelpClass.ConvertDateType(StrDate, "-"))); //기간(FR)
                dataPackFormula.Add(new PSH_DataPackClass("@EndDate", EndDate == "21001231" ? "All" : dataHelpClass.ConvertDateType(EndDate, "-"))); //기간(TO)
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLID, ""))); //사업장명

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@StrDate", StrDate == "19000101" ? "1900-01-01" : dataHelpClass.ConvertDateType(StrDate, "-"))); //기간(FR)
                dataPackParameter.Add(new PSH_DataPackClass("@EndDate", EndDate == "21001231" ? "2100-12-31" : dataHelpClass.ConvertDateType(EndDate, "-"))); //기간(TO)
                dataPackParameter.Add(new PSH_DataPackClass("@ItmsGrpCod", ItmsGrpCod)); //품목그룹
                dataPackParameter.Add(new PSH_DataPackClass("@SItemCode", SItemCode)); //품목코드(S)
                dataPackParameter.Add(new PSH_DataPackClass("@EItemCode", EItemCode)); //품목코드(E)
                dataPackParameter.Add(new PSH_DataPackClass("@SCardCode", SCardCode)); //거래처(S)
                dataPackParameter.Add(new PSH_DataPackClass("@ECardCode", ECardCode)); //거래처(E)
                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@ItmBsort", ItmBsort)); //대분류(코드)
                dataPackParameter.Add(new PSH_DataPackClass("@ItmMsort", ItmMsort)); //중분류(코드)    
                dataPackParameter.Add(new PSH_DataPackClass("@OrdType", OrdTyp)); //품의구분
                dataPackParameter.Add(new PSH_DataPackClass("@POType", POType)); //품의형태
                dataPackParameter.Add(new PSH_DataPackClass("@DocType", DocType)); //품목/서비스
                dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode)); //공정
                dataPackParameter.Add(new PSH_DataPackClass("@APYN", APYN)); //AP송장여부

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 공정코드 콤보박스 바인딩
        /// </summary>
        private void PS_MM204_SetCpCode()
        {
            string sQry;
            string stdYear;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                stdYear = oForm.DataSources.UserDataSources.Item("StrDate").Value.ToString().Trim().Replace(".", "").Substring(0, 4);

                //기존 콤보 데이터 삭제
                if (oForm.Items.Item("CpCode").Specific.ValidValues.Count > 0)
                {
                    for (int i = oForm.Items.Item("CpCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oForm.Items.Item("CpCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }

                oForm.Items.Item("CpCode").Specific.ValidValues.Add("%", "전체");

                sQry = "  SELECT	T1.U_ProcCode,";
                sQry += "           T2.U_CpName";
                sQry += " FROM      [@PS_MM030H] AS T0";
                sQry += "           INNER JOIN";
                sQry += "           [@PS_MM030L] AS T1";
                sQry += "               ON T0.DocEntry = T1.DocEntry";
                sQry += "           LEFT JOIN";
                sQry += "           [@PS_PP001L] AS T2";
                sQry += "               ON T1.U_ProcCode = T2.U_CpCode";
                sQry += " WHERE     T0.U_BPLId = '" + oForm.DataSources.UserDataSources.Item("BPLId").Value.ToString().Trim() + "'";
                sQry += "           AND CONVERT(VARCHAR(4), T0.U_DocDate, 112) = '" + stdYear + "'";
                sQry += "           AND ISNULL(T1.U_ProcCode, '') <> ''";
                sQry += "           AND ISNULL(T1.U_ProcName, '') <> ''";
                sQry += " GROUP BY  T1.U_ProcCode,";
                sQry += "           T2.U_CpName";
                sQry += " ORDER BY  T1.U_ProcCode";

                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("CpCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("CpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01") //출력버튼
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM204_PrintReport);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "SItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SItemCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "EItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("EItemCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "SCardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SCardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ECardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ECardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItmBsort")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItmMsort")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "BPLId")
                    {
                        PS_MM204_SetCpCode();
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

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "ItmMsort")
                        {
                            PS_MM204_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                        else if (pVal.ItemUID == "ItmBsort")
                        {
                            PS_MM204_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                        else if (pVal.ItemUID == "StrDate" || pVal.ItemUID == "EndDate")
                        {
                            PS_MM204_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                            break;
                        case "1285": //복원
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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
                        case "1285": //복원
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
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
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
