using System;
using System.Collections.Generic;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합재무제표
	/// </summary>
	internal class PS_CO600 : PSH_BaseClass
	{

		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;

		public SAPbouiCOM.Grid oGrid01;
		public SAPbouiCOM.Grid oGrid02;
		public SAPbouiCOM.Grid oGrid03;
		public SAPbouiCOM.Grid oGrid04;

		public SAPbouiCOM.DataTable oDS_PS_CO600A;
		public SAPbouiCOM.DataTable oDS_PS_CO600B;
		public SAPbouiCOM.DataTable oDS_PS_CO600C;
		public SAPbouiCOM.DataTable oDS_PS_CO600D;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO600.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO600_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO600");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

				oForm.Freeze(true);
                PS_CO600_CreateItems();
                PS_CO600_ComboBox_Setting();

                oForm.Items.Item("FrDt01").Specific.Value = DateTime.Now.ToString("yyyy0101"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY0101");
				oForm.Items.Item("ToDt01").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
				
				oForm.Items.Item("Folder01").Specific.Select();				
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

        private void PS_CO600_CreateItems()
        {
            try
            {
                //oForm.Freeze(true);

                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oGrid02 = oForm.Items.Item("Grid02").Specific;
                oGrid03 = oForm.Items.Item("Grid03").Specific;
                oGrid04 = oForm.Items.Item("Grid04").Specific;

                oForm.DataSources.DataTables.Add("PS_CO600A");
                oForm.DataSources.DataTables.Add("PS_CO600B");
                oForm.DataSources.DataTables.Add("PS_CO600C");
                oForm.DataSources.DataTables.Add("PS_CO600D");

                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_CO600A");
                oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_CO600B");
                oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_CO600C");
                oGrid04.DataTable = oForm.DataSources.DataTables.Item("PS_CO600D");

                oDS_PS_CO600A = oForm.DataSources.DataTables.Item("PS_CO600A");
                oDS_PS_CO600B = oForm.DataSources.DataTables.Item("PS_CO600B");
                oDS_PS_CO600C = oForm.DataSources.DataTables.Item("PS_CO600C");
                oDS_PS_CO600D = oForm.DataSources.DataTables.Item("PS_CO600D");

                //조회기간(시작)
                oForm.DataSources.UserDataSources.Add("FrDt01", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDt01").Specific.DataBind.SetBound(true, "", "FrDt01");

                //조회기간(종료)
                oForm.DataSources.UserDataSources.Add("ToDt01", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDt01").Specific.DataBind.SetBound(true, "", "ToDt01");

                //출력구분
                oForm.DataSources.UserDataSources.Add("Ctgr01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Ctgr01").Specific.DataBind.SetBound(true, "", "Ctgr01");

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                //oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO600_ComboBox_Setting()
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("Ctgr01").Specific.ValidValues.Add("10", "K-GAAP");
                oForm.Items.Item("Ctgr01").Specific.ValidValues.Add("20", "K-IFRS");
                oForm.Items.Item("Ctgr01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// ChooseFromList
        /// </summary>
        private void PS_CO600_CF_ChooseFromList()
        {
            try
            {
                ////ChooseFromList 설정
                //    Dim oCFLs               As SAPbouiCOM.ChooseFromListCollection
                //    Dim oCons               As SAPbouiCOM.Conditions
                //    Dim oCon                As SAPbouiCOM.Condition
                //    Dim oCFL                As SAPbouiCOM.ChooseFromList
                //    Dim oCFLCreationParams  As SAPbouiCOM.ChooseFromListCreationParams
                //    Dim oEdit               As SAPbouiCOM.EditText
                //    Dim oColumn             As SAPbouiCOM.Column
                //
                //    Set oEdit = oForm01.Items("ItemCode").Specific
                //    Set oCFLs = oForm01.ChooseFromLists
                //    Set oCFLCreationParams = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                //
                //    oCFLCreationParams.ObjectType = "4"
                //    oCFLCreationParams.uniqueID = "CFLITEMCD"
                //    oCFLCreationParams.MultiSelection = False
                //    Set oCFL = oCFLs.Add(oCFLCreationParams)
                //
                //'    Set oCons = oCFL.GetConditions()
                //'    Set oCon = oCons.Add()
                //'    oCon.Alias = "CardType"
                //'    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                //'    oCon.CondVal = "C"
                //'    oCFL.SetConditions oCons
                //
                //    oEdit.ChooseFromListUID = "CFLITEMCD"
                //    oEdit.ChooseFromListAlias = "ItemCode"
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormItemEnabled
        /// </summary>
        private void PS_CO600_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
        /// 
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO600_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                //if (RowIserted == false) //행추가여부
                //{
                //    oDS_PS_CO600L.InsertRecord(oRow);
                //}

                //oMat01.AddRow();
                //oDS_PS_CO600L.Offset = oRow;
                //oDS_PS_CO600L.setValue("U_LineNum", oRow, oRow + 1);
                //oMat01.LoadFromDataSource();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PS_CO600_DataValidCheck()
        {
            bool functionReturnValue = false;

            try
            {

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            
            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_CO600_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            try
            {
                switch (oUID)
                {
                    // Case "CntcCode01"

                    //     If Trim(oForm01.Items("CntcCode01").Specific.VALUE) = "9999999" Then
                    //         oForm01.Items("CntcName01").Specific.VALUE = "공용" '성명
                    //     Else
                    //         oForm01.Items("CntcName01").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode01").Specific.VALUE) & "'") '성명
                    //     End If

                    // Case "TeamCode01"

                    //     If Trim(oForm01.Items("TeamCode01").Specific.VALUE) = oForm01.Items("BPLID01").Specific.Selected.VALUE & "999" Then
                    //         oForm01.Items("TeamName01").Specific.VALUE = "전체공용"
                    //     ElseIf Trim(oForm01.Items("TeamCode01").Specific.VALUE) = "Z" & oForm01.Items("BPLID01").Specific.Selected.VALUE & "99" Then
                    //         oForm01.Items("TeamName01").Specific.VALUE = "사용부서없음"
                    //     Else
                    //         oForm01.Items("TeamName01").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & Trim(oForm01.Items("TeamCode01").Specific.VALUE) & "'", " AND Code = '1'") '팀
                    //     End If

                    // Case "CntcCode02"

                    //     If Trim(oForm01.Items("CntcCode02").Specific.VALUE) = "9999999" Then
                    //         oForm01.Items("CntcName02").Specific.VALUE = "공용" '성명
                    //     Else
                    //         oForm01.Items("CntcName02").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode02").Specific.VALUE) & "'") '성명
                    //     End If

                    //  Case "CntcCode03"

                    //     If Trim(oForm01.Items("CntcCode03").Specific.VALUE) = "9999999" Then
                    //         oForm01.Items("CntcName03").Specific.VALUE = "공용" '성명
                    //     Else
                    //         oForm01.Items("CntcName03").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode03").Specific.VALUE) & "'") '성명
                    //     End If

                    // Case "CntcCode05"

                    //     If Trim(oForm01.Items("CntcCode05").Specific.VALUE) = "9999999" Then
                    //         oForm01.Items("CntcName05").Specific.VALUE = "공용" '성명
                    //     Else
                    //         oForm01.Items("CntcName05").Specific.VALUE = MDC_GetData.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" & Trim(oForm01.Items("CntcCode05").Specific.VALUE) & "'") '성명
                    //     End If

                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 그리드 출력
        /// </summary>
        /// <param name="pGrid">B1 그리드 객체</param>
        /// <param name="pDataTable">B1 데이터테이블 객체</param>
        private void PS_CO600_MTX(SAPbouiCOM.Grid pGrid, SAPbouiCOM.DataTable pDataTable)
        {
            string Query01 = string.Empty;
            string FrDt = string.Empty; //조회기간(시작)
            string ToDt = string.Empty; //조회기간(종료)
            string Ctgr = string.Empty; //출력구분
            string PrtCls = string.Empty; //그리드, 리포트 출력구분

            string errCode = string.Empty;

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

            try
            {
                oForm.Freeze(true);

                FrDt = oForm.Items.Item("FrDt01").Specific.Value.ToString().Trim(); //조회기간(시작)
                ToDt = oForm.Items.Item("ToDt01").Specific.Value.ToString().Trim(); //조회기간(종료)
                Ctgr = oForm.Items.Item("Ctgr01").Specific.Selected.Value.ToString().Trim(); //출력구분
                PrtCls = "G"; //그리드출력

                if (pGrid.Item.UniqueID == "Grid01") //대차대조표
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        Query01 = "         EXEC PS_CO600_01 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    } //K-IFRS
                    else
                    {
                        Query01 = "         EXEC PS_CO600_21 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    }
                }
                else if (pGrid.Item.UniqueID == "Grid02") //제조원가명세서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        Query01 = "         EXEC PS_CO600_02 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    } //K-IFRS
                    else
                    {
                        Query01 = "         EXEC PS_CO600_22 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    }
                }
                else if (pGrid.Item.UniqueID == "Grid03") //매출원가명세서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        Query01 = "         EXEC PS_CO600_03 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    } //K-IFRS
                    else
                    {
                        Query01 = "         EXEC PS_CO600_23 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    }
                }
                else if (pGrid.Item.UniqueID == "Grid04") //손익계산서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        Query01 = "         EXEC PS_CO600_04 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    } //K-IFRS
                    else
                    {
                        Query01 = "         EXEC PS_CO600_24 '";
                        Query01 = Query01 + FrDt + "','";
                        Query01 = Query01 + ToDt + "','";
                        Query01 = Query01 + PrtCls + "'";
                    }
                }

                pGrid.DataTable.Clear();
                pDataTable.ExecuteQuery(Query01);

                pGrid.Columns.Item(2).RightJustified = true;
                pGrid.Columns.Item(3).RightJustified = true;
                pGrid.Columns.Item(4).RightJustified = true;
                pGrid.Columns.Item(5).RightJustified = true;
                pGrid.Columns.Item(6).RightJustified = true;

                if (pGrid.Rows.Count == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                pGrid.AutoResizeColumns();
            }
            catch (Exception ex)
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
                oForm.Update();
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_CO600_FormResize()
        {
            try
            {
                //그룹박스 크기 동적 할당
                oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 60;
                oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

                if (oGrid01.Columns.Count > 0)
                {
                    oGrid01.AutoResizeColumns();
                }

                if (oGrid03.Columns.Count > 0)
                {
                    oGrid03.AutoResizeColumns();
                }

                if (oGrid03.Columns.Count > 0)
                {
                    oGrid03.AutoResizeColumns();
                }

                if (oGrid04.Columns.Count > 0)
                {
                    oGrid04.AutoResizeColumns();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        /// <param name="pButtonID">재무제표 출력 버튼 ID</param>
        [STAThread]
        private void PS_CO600_Print_Report(object pButtonID)
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            System.DateTime FrDt; ; //조회기간(시작)
            System.DateTime ToDt; ; //조회기간(종료)
            string Ctgr; //출력구분
            string PrtCls; //그리드, 리포트 출력구분

            FrDt = DateTime.ParseExact(oForm.Items.Item("FrDt01").Specific.Value, "yyyyMMdd", null); //조회기간(시작)
            ToDt = DateTime.ParseExact(oForm.Items.Item("ToDt01").Specific.Value, "yyyyMMdd", null); //조회기간(종료)
            Ctgr = oForm.Items.Item("Ctgr01").Specific.Selected.Value.ToString().Trim(); //출력구분
            PrtCls = "R"; //리포트출력

            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (pButtonID.ToString() == "BtnPrt01") //대차대조표
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        ReportName = "PS_CO600_51.rpt"; //프로시저 : PS_CO600_01
                        WinTitle = "[PS_CO600] 대차대조표";
                    }
                    else //K-IFRS
                    {
                        ReportName = "PS_CO600_61.rpt"; //프로시저 : PS_CO600_21(구현안됨)
                        WinTitle = "[PS_CO600] 재무상태표";
                    }
                }
                else if (pButtonID.ToString() == "BtnPrt02") //제조원가명세서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        ReportName = "PS_CO600_52.rpt"; //프로시저 : PS_CO600_02
                        WinTitle = "[PS_CO600] 제조원가명세서";
                    }
                    else //K-IFRS
                    {
                        ReportName = "PS_CO600_62.rpt"; //프로시저 : PS_CO600_22(구현안됨)
                        WinTitle = "[PS_CO600] 제조원가명세서";
                    }
                }
                else if (pButtonID.ToString() == "BtnPrt03") //매출원가명세서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        ReportName = "PS_CO600_53.rpt"; //프로시저 : PS_CO600_03
                        WinTitle = "[PS_CO600] 매출원가명세서";
                    } //K-IFRS
                    else
                    {
                        ReportName = "PS_CO600_63.rpt"; //프로시저 : PS_CO600_23(구현안됨)
                        WinTitle = "[PS_CO600] 매출원가명세서";
                    }
                }
                else if (pButtonID.ToString() == "BtnPrt04") //손익계산서
                {
                    if (Ctgr == "10") //K-GAAP
                    {
                        ReportName = "PS_CO600_54.rpt"; //프로시저 : PS_CO600_04
                        WinTitle = "[PS_CO600] 손익계산서";
                    } //K-IFRS
                    else
                    {
                        ReportName = "PS_CO600_64.rpt"; //프로시저 : PS_CO600_24(구현안됨)
                        WinTitle = "[PS_CO600] 포괄손익계산서";
                    }
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt)); //조회일자(시작)
                dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt)); //조회일자(종료)
                dataPackParameter.Add(new PSH_DataPackClass("@PrtCls", PrtCls)); //그리드, 리포트 출력구분

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
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
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "BtnSrch01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_CO600_MTX(oGrid01, oDS_PS_CO600A); //대차대조표(재무상태표)
                            PS_CO600_MTX(oGrid02, oDS_PS_CO600B); //제조원가명세서
                            PS_CO600_MTX(oGrid03, oDS_PS_CO600C); //매출원가명세서
                            PS_CO600_MTX(oGrid04, oDS_PS_CO600D); //손익계산서(포괄손익계산서)

                            //PS_CO600_MTX01(); //대차대조표(재무상태표)
                            //PS_CO600_MTX02(); //제조원가명세서
                            //PS_CO600_MTX03(); //매출원가명세서
                            //PS_CO600_MTX04(); //손익계산서(포괄손익계산서)
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrt01" || pVal.ItemUID == "BtnPrt02" || pVal.ItemUID == "BtnPrt03" || pVal.ItemUID == "BtnPrt04") //리포트
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(PS_CO600_Print_Report));
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start(pVal.ItemUID);
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
                    //폴더를 사용할 때는 필수 소스_S
                    if (pVal.ItemUID == "Folder01") //Folder01이 선택되었을 때
                    {
                        oForm.PaneLevel = 1;
                    }

                    if (pVal.ItemUID == "Folder02") //Folder02가 선택되었을 때
                    {
                        oForm.PaneLevel = 2;
                    }

                    if (pVal.ItemUID == "Folder03") //Folder03가 선택되었을 때
                    {
                        oForm.PaneLevel = 3;
                    }
                    
                    if (pVal.ItemUID == "Folder04") //Folder04가 선택되었을 때
                    {
                        oForm.PaneLevel = 4;
                    }
                    //폴더를 사용할 때는 필수 소스_E

                    if (pVal.ItemUID == "PS_CO600")
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
                    PS_CO600_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row > 0)
                            {
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row == -1)
                        {
                        }
                        else
                        {
                            if (oGrid01.Rows.SelectedRows.Count > 0)
                            {
                            }
                            else
                            {
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO600_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO600_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid04);
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO600_FormResize();
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
                        //Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
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
                        //Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
                        case "1281": //찾기
                            break;
                            //Call PS_CO600_FormItemEnabled '//UDO방식
                        case "1282": //추가
                            break;
                            //Call PS_CO600_FormItemEnabled '//UDO방식
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
