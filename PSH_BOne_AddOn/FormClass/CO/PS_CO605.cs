using System;
using System.Collections.Generic;

using SAPbouiCOM;

using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합수불부
	/// </summary>
	internal class PS_CO605 : PSH_BaseClass
	{
		private string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		private SAPbouiCOM.Grid oGrid01;

		private SAPbouiCOM.DataTable oDS_PS_CO605A;

		//public SAPbouiCOM.Form oBaseForm01; //부모폼
		//public string oBaseItemUID01;
		//public string oBaseColUID01;
		//public int oBaseColRow01;
		//public string oBaseTradeType01;
		//public string oBaseItmBsort01;
			
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO605.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO605_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO605");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

				oForm.Freeze(true);
                PS_CO605_CreateItems();
                PS_CO605_ComboBox_Setting();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_CO605_CreateItems()
        {
            SAPbouiCOM.CheckBox oChkBox = null;

            try
            {
                oForm.Freeze(true);
                
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                //oGrid01.SelectionMode = ms_NotSupported

                oForm.DataSources.DataTables.Add("PS_CO605A");
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_CO605A");
                oDS_PS_CO605A = oForm.DataSources.DataTables.Item("PS_CO605A");

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

                //전기일자(Fr)
                oForm.DataSources.UserDataSources.Add("StrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("StrDate").Specific.DataBind.SetBound(true, "", "StrDate");
                oForm.DataSources.UserDataSources.Item("StrDate").Value = DateTime.Now.ToString("yyyyMM01"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM01");

                //전기일자(To)
                oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");
                oForm.DataSources.UserDataSources.Item("EndDate").Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

                //재고계정
                oForm.DataSources.UserDataSources.Add("AcctCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("AcctCode").Specific.DataBind.SetBound(true, "", "AcctCode");

                //창고
                oForm.DataSources.UserDataSources.Add("WhsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("WhsCode").Specific.DataBind.SetBound(true, "", "WhsCode");

                //대분류
                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                //중분류
                oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

                //출력구분
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Gubun").Specific.DataBind.SetBound(true, "", "Gubun");

                //체크박스 처리
                oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                oChkBox = oForm.Items.Item("Check01").Specific;
                oChkBox.ValOn = "Y";
                oChkBox.ValOff = "N";
                oChkBox.DataBind.SetBound(true, "", "Check01");

                oForm.DataSources.UserDataSources.Item("Check01").Value = "N"; //미체크로 값을 주고 폼을 로드

                oForm.DataSources.UserDataSources.Add("Check02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                oChkBox = oForm.Items.Item("Check02").Specific;
                oChkBox.ValOn = "Y";
                oChkBox.ValOff = "N";
                oChkBox.DataBind.SetBound(true, "", "Check02");

                oForm.DataSources.UserDataSources.Item("Check02").Value = "N"; //미체크로 값을 주고 폼을 로드

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oChkBox);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO605_ComboBox_Setting()
        {
        
            SAPbouiCOM.ComboBox oCombo = null;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                //콤보에 기본값설정

                //사업장
                oCombo = oForm.Items.Item("BPLID").Specific;
                oCombo.ValidValues.Add("%", "전체");
                oCombo.ValidValues.Add("1", "창원사업장");
                oCombo.ValidValues.Add("2", "부산사업장");
                oCombo.ValidValues.Add("6", "안강+울산사업장");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //재고계정
                oCombo = oForm.Items.Item("AcctCode").Specific;
                oCombo.ValidValues.Add("11506100", "원재료");
                oCombo.ValidValues.Add("11502100", "제품");
                oCombo.ValidValues.Add("11501100", "상품");
                oCombo.ValidValues.Add("11507100", "저장품");
                oCombo.ValidValues.Add("11503100", "재공품");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //창고
                oCombo = oForm.Items.Item("WhsCode").Specific;
                sQry = "SELECT WhsCode, WhsName From OWHS";
                oRecordSet01.DoQuery(sQry);
                oCombo.ValidValues.Add("000", "전체");
                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //대분류
                oCombo = oForm.Items.Item("ItmBSort").Specific;
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
                oRecordSet01.DoQuery(sQry);
                oCombo.ValidValues.Add("001", "전체");
                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //중분류
                oCombo = oForm.Items.Item("ItmMSort").Specific;
                sQry = "SELECT U_Code,U_CodeName FROM [@PSH_ITMMSORT] Order by U_Code";
                oRecordSet01.DoQuery(sQry);

                if (oForm.Items.Item("ItmMSort").Specific.ValidValues.Count == 0)
                {
                    oCombo.ValidValues.Add("00001", "전체");
                }

                while (!oRecordSet01.EoF)
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //출력구분
                oCombo = oForm.Items.Item("Gubun").Specific;
                oCombo.ValidValues.Add("1", "개별");
                oCombo.ValidValues.Add("2", "집계");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO605_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //각모드에따른 아이템설정
                    //PS_CO605_FormClear //UDO방식
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //각모드에따른 아이템설정
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //각모드에따른 아이템설정
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO605_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                //if (RowIserted = false) //행추가여부
                //{
                //    oDS_PS_CO605L.InsertRecord(oRow);
                //}
                    
                //oMat01.AddRow();
                //oDS_PS_CO605L.Offset = oRow;
                //oDS_PS_CO605L.setValue("U_LineNum", oRow, oRow + 1);
                //oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_CO605_DataValidCheck()
        {
            bool returnValue = false;

            try
            {
                returnValue = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 수불부조회
        /// </summary>
        private void PS_CO605_MTX01()
        {
            string Query01;
            string ItmBsort;
            string ItmMsort;
            string BPLId;
            string StrDate;
            string EndDate;
            string AcctCode;
            string WhsCode;
            string ChkBox;
            string ChkBox02;
            string Gubun;

            string errCode = string.Empty;

            //SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

            //조회조건
            ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
            ItmMsort = oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();
            BPLId = oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim();
            StrDate = oForm.Items.Item("StrDate").Specific.Value.ToString().Trim();
            EndDate = oForm.Items.Item("EndDate").Specific.Value.ToString().Trim();
            AcctCode = oForm.Items.Item("AcctCode").Specific.Selected.Value.ToString().Trim();
            WhsCode = oForm.Items.Item("WhsCode").Specific.Selected.Value.ToString().Trim();
            ChkBox = oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim();
            ChkBox02 = oForm.DataSources.UserDataSources.Item("Check02").Value.ToString().Trim();
            Gubun = oForm.Items.Item("Gubun").Specific.Selected.Value.ToString().Trim();

            try
            {
                oForm.Freeze(true);

                if (string.IsNullOrEmpty(StrDate))
                {
                    StrDate = "19000101";
                }
                    
                if (string.IsNullOrEmpty(EndDate))
                {
                    EndDate = "21001231";
                }
                    
                if (ItmBsort == "001")
                {
                    ItmBsort = "%";
                }
                    
                if (ItmMsort == "00001")
                {
                    ItmMsort = "%";
                }

                if (Gubun == "1") //수불개별
                {
                    if (ChkBox02 == "Y") //포장사업팀용
                    {
                        Query01 = "EXEC [PS_MM209_10] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "'";
                    }
                    else
                    {
                        Query01 = "EXEC [PS_MM209_02] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_CO605_02'";
                    }
                }
                else //수불집계
                {
                    Query01 = "EXEC [PS_MM209_04] '" + BPLId + "','" + StrDate + "','" + EndDate + "','" + AcctCode + "', '" + WhsCode + "', '" + ChkBox + "', '" + ItmBsort + "', '" + ItmMsort + "', 'PS_CO605_02'";
                }

                oGrid01.DataTable.Clear();
                oDS_PS_CO605A.ExecuteQuery(Query01);
                //oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

                if (oGrid01.Rows.Count == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                oGrid01.Columns.Item(5).RightJustified = true;
                oGrid01.Columns.Item(6).RightJustified = true;
                oGrid01.Columns.Item(7).RightJustified = true;
                oGrid01.Columns.Item(8).RightJustified = true;
                oGrid01.Columns.Item(9).RightJustified = true;
                oGrid01.Columns.Item(10).RightJustified = true;
                oGrid01.Columns.Item(11).RightJustified = true;
                oGrid01.Columns.Item(12).RightJustified = true;
                oGrid01.Columns.Item(13).RightJustified = true;
                oGrid01.Columns.Item(14).RightJustified = true;
                oGrid01.Columns.Item(15).RightJustified = true;

                if (Gubun == "1")
                {
                    oGrid01.Columns.Item(16).RightJustified = true;
                }

                oGrid01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
                ProgBar01.Value = 100;
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_CO605_FormResize()
        {
            try
            {
                if (oGrid01.Columns.Count > 0)
                {
                    oGrid01.AutoResizeColumns();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        [STAThread]
        private void PS_CO605_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;
            string sQry;
            
            string ItmBsort;
            string ItmMsort;
            string BPLId;
            string StrDate;
            string EndDate;
            string AcctCode;
            string WhsCode;
            string ChkBox;
            //string ChkBox02;
            string Gubun;

            string BPLName; //Formula 전달용

            //조회조건문
            ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
            ItmMsort = oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();
            BPLId = oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim();
            StrDate = oForm.Items.Item("StrDate").Specific.Value.ToString().Trim();
            EndDate = oForm.Items.Item("EndDate").Specific.Value.ToString().Trim();
            AcctCode = oForm.Items.Item("AcctCode").Specific.Selected.Value.ToString().Trim();
            WhsCode = oForm.Items.Item("WhsCode").Specific.Selected.Value.ToString().Trim();
            ChkBox = oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim();
            //ChkBox02 = oForm.DataSources.UserDataSources.Item("Check02").Value.ToString().Trim();
            Gubun = oForm.Items.Item("Gubun").Specific.Selected.Value.ToString().Trim();

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (string.IsNullOrEmpty(StrDate))
                {
                    StrDate = "19000101";
                }

                if (string.IsNullOrEmpty(EndDate))
                {
                    EndDate = "21001231";
                }

                if (ItmBsort == "001")
                {
                    ItmBsort = "%";
                }

                if (ItmMsort == "00001")
                {
                    ItmMsort = "%";
                }

                if (Gubun == "1")
                {
                    WinTitle = "[PS_CO605] 수불명세서";
                    ReportName = "PS_CO605_01.RPT";
                }
                else if (Gubun == "2")
                {
                    WinTitle = "[PS_CO605] 수불명세서(집계)";
                    ReportName = "PS_CO605_02.RPT";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@StrDate", (StrDate == "" ? "All" : dataHelpClass.ConvertDateType(StrDate, "-"))));
                dataPackFormula.Add(new PSH_DataPackClass("@EndDate", (EndDate == "" ? "All" : dataHelpClass.ConvertDateType(EndDate, "-"))));
                if (BPLId == "6")
                {
                    BPLName = "안강+울산 사업장";
                }
                else if (BPLId == "%")
                {
                    BPLName = "통합사업장";
                }
                else
                {
                    sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLId + "'";
                    oRecordSet.DoQuery(sQry);
                    BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                }
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLName));
                dataPackFormula.Add(new PSH_DataPackClass("@AcctCode", AcctCode));
                dataPackFormula.Add(new PSH_DataPackClass("@ChkBox", ChkBox));

                sQry = "SELECT WhsName From OWHS where WhsCode = '" + WhsCode + "'";
                oRecordSet.DoQuery(sQry);
                dataPackFormula.Add(new PSH_DataPackClass("@WhsName", (WhsCode == "000" ? "전체" : oRecordSet.Fields.Item(0).Value)));

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId)); //BPLId
                dataPackParameter.Add(new PSH_DataPackClass("@FrDate", dataHelpClass.ConvertDateType(StrDate, "-"))); //FrDate
                dataPackParameter.Add(new PSH_DataPackClass("@ToDate", dataHelpClass.ConvertDateType(EndDate, "-"))); //ToDate
                dataPackParameter.Add(new PSH_DataPackClass("@AcctCode", AcctCode)); //AcctCode
                dataPackParameter.Add(new PSH_DataPackClass("@WareHouse", WhsCode)); //WareHouse
                dataPackParameter.Add(new PSH_DataPackClass("@Wgt", ChkBox)); //Wgt
                dataPackParameter.Add(new PSH_DataPackClass("@ItmBsort", ItmBsort)); //ItmBsort
                dataPackParameter.Add(new PSH_DataPackClass("@ItmMsort", ItmMsort)); //ItmMsort
                dataPackParameter.Add(new PSH_DataPackClass("@Class", "PS_MM209")); //Class

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch(Exception ex)
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_CO605_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                PS_CO605_MTX01();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrt")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_CO605_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }

                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PS_CO605")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO605_FormResize();
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
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
                if (BusinessObjectInfo.BeforeAction == true)
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
                else if (BusinessObjectInfo.BeforeAction == false)
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
