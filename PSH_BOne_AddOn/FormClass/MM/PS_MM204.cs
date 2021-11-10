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
		private SAPbouiCOM.DBDataSource oDS_PS_MM204H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM204L; //등록라인
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

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

                oForm.Items.Item("StrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

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

                oForm.Items.Item("Rad03").Specific.GroupWith("Rad01");
                oForm.Items.Item("Rad03").Specific.GroupWith("Rad02");
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
                StrDate = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("StrDate").Value.ToString().Trim()) ? "19000101" : oForm.DataSources.UserDataSources.Item("StrDate").Value.ToString().Trim(); //기간(FR)
                EndDate = string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("EndDate").Value.ToString().Trim()) ? "21001231" : oForm.DataSources.UserDataSources.Item("EndDate").Value.ToString().Trim(); //기간(TO)
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
                //프로시저 : PS_MM204_00

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                //Formula                
                dataPackFormula.Add(new PSH_DataPackClass("@StrDate", StrDate == "19000101" ? "All" : dataHelpClass.ConvertDateType(StrDate, "-"))); //기간(FR)
                dataPackFormula.Add(new PSH_DataPackClass("@EndDate", EndDate == "21001231" ? "All" : dataHelpClass.ConvertDateType(EndDate, "-"))); //기간(TO)
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLID, ""))); //사업장명

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@StdDate", StrDate)); //기간(FR)
                dataPackParameter.Add(new PSH_DataPackClass("@EndDate", EndDate)); //기간(TO)
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

        #region Raise_ItemEvent
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			int ErrNum = 0;
        //			SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						if (pval.ItemUID == "1") {
        //							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //								//                        If PS_MM204_DelHeaderSpaceLine = False Then
        //								//                            BubbleEvent = False
        //								//                            Exit Sub
        //								//                        End If
        //								//                        If MatrixSpaceLineDel = False Then
        //								//                            BubbleEvent = False
        //								//                            Exit Sub
        //								//                        End If
        //							}

        //						//출력버튼 클릭시
        //						} else if (pval.ItemUID == "Btn01") {
        //							if (PS_MM204_DelHeaderSpaceLine() == false) {
        //								BubbleEvent = false;
        //								return;
        //							} else {
        //								PS_MM204_PrintReport();
        //							}
        //						}
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						if (pval.CharPressed == 9) {
        //							////헤더
        //							if (pval.ItemUID == "SItemCode") {
        //								//UPGRADE_WARNING: oForm.Items(SItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("SItemCode").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							if (pval.ItemUID == "EItemCode") {
        //								//UPGRADE_WARNING: oForm.Items(EItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("EItemCode").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							if (pval.ItemUID == "SCardCode") {
        //								//UPGRADE_WARNING: oForm.Items(SCardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("SCardCode").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							if (pval.ItemUID == "ECardCode") {
        //								//UPGRADE_WARNING: oForm.Items(ECardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("ECardCode").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							if (pval.ItemUID == "U_ItmBsort") {
        //								//UPGRADE_WARNING: oForm.Items(U_ItmBsort).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("U_ItmBsort").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							if (pval.ItemUID == "ItmMsort") {
        //								//UPGRADE_WARNING: oForm.Items(ItmMsort).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.Value)) {
        //									SubMain.Sbo_Application.ActivateMenuItem(("7425"));
        //									BubbleEvent = false;
        //								}
        //							}
        //							////라인
        //							//                    If pval.ItemUID = "Mat01" Then
        //							//                        If pval.ColUID = "PP070No" Then
        //							//                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
        //							//                                Sbo_Application.ActivateMenuItem ("7425")
        //							//                                BubbleEvent = False
        //							//                            End If
        //							//                        End If
        //							//                    End If
        //						}
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //						////5
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CLICK:
        //						////6
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //						////7
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //						////8
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //						////10

        //						if (pval.ItemChanged == true) {
        //							if (pval.ItemUID == "ItmMsort") {
        //								PS_MM204_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
        //							}

        //							if (pval.ItemUID == "U_ItmBsort") {
        //								PS_MM204_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);
        //							}
        //						}
        //						break;


        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //						////11
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //						////18
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //						////19
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //						////20
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //						////27
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //						////3
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						break;
        //				}

        //				//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						break;
        //					//                If pval.ItemUID = "1" Then
        //					//                    If oForm.Mode = fm_ADD_MODE Then
        //					//                        oForm.Mode = fm_OK_MODE
        //					//                        Call Sbo_Application.ActivateMenuItem("1282")
        //					//                    ElseIf oForm.Mode = fm_OK_MODE Then
        //					//                        FormItemEnabled
        //					//                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
        //					//                    End If
        //					//                End If
        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //						////5
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CLICK:
        //						////6
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //						////7
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //						////8
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //						////10
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //						////11
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //						////18
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //						////19
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //						////20
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //						////27
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //						////3
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						SubMain.RemoveForms(oFormUniqueID);
        //						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oForm = null;
        //						break;
        //					//                Set oMat01 = Nothing
        //				}
        //			}
        //			return;
        //			Raise_ItemEvent_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			if (ErrNum == 101) {
        //				ErrNum = 0;
        //				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //				BubbleEvent = false;
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region Raise_MenuEvent
        //		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						break;
        //					case "1285":
        //						//복원
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}

        //				//-----------------------------------------------------------------------------------------------------------
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1285":
        //						//복원
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if ((eventInfo.BeforeAction == true)) {

        //			} else if ((eventInfo.BeforeAction == false)) {
        //				////작업
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion


    }
}
