using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 비근무일수현황
    /// </summary>
    internal class PS_PP750 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid mainGrid;
        private SAPbouiCOM.DataTable oDS_PS_PP750;
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP750.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP750_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP750");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PP750_CreateItems();
                PS_PP750_SetComboBox();
                PS_PP750_InitializeForm();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <returns></returns>
        private void PS_PP750_CreateItems()
        {
            try
            {
                mainGrid = oForm.Items.Item("mainGrid").Specific;
                oForm.DataSources.DataTables.Add("PS_PP750");

                mainGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP750");
                oDS_PS_PP750 = oForm.DataSources.DataTables.Item("PS_PP750");

                //수주일자(FR)
                oForm.DataSources.UserDataSources.Add("FrDocDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDocDt").Specific.DataBind.SetBound(true, "", "FrDocDt");

                //수주일자(TO)
                oForm.DataSources.UserDataSources.Add("ToDocDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDocDt").Specific.DataBind.SetBound(true, "", "ToDocDt");

                //납기일자(FR)
                oForm.DataSources.UserDataSources.Add("FrDueDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDueDt").Specific.DataBind.SetBound(true, "", "FrDueDt");

                //납기일자(TO)
                oForm.DataSources.UserDataSources.Add("ToDueDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDueDt").Specific.DataBind.SetBound(true, "", "ToDueDt");

                //생산담당(사번)
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //생산담당(성명)
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //자체/외주
                oForm.DataSources.UserDataSources.Add("InOut", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("InOut").Specific.DataBind.SetBound(true, "", "InOut");

                //장비/공구
                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                //거래처구분
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

                //조회구분
                oForm.DataSources.UserDataSources.Add("SrchType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("SrchType").Specific.DataBind.SetBound(true, "", "SrchType");

                //작번(품목코드)
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //품목규격
                oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

                //품의여부
                oForm.DataSources.UserDataSources.Add("MM030YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("MM030YN").Specific.DataBind.SetBound(true, "", "MM030YN");

                //입고여부
                oForm.DataSources.UserDataSources.Add("MM050YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("MM050YN").Specific.DataBind.SetBound(true, "", "MM050YN");

                //검수여부
                oForm.DataSources.UserDataSources.Add("MM070YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("MM070YN").Specific.DataBind.SetBound(true, "", "MM070YN");

                //품의구분
                oForm.DataSources.UserDataSources.Add("OrdType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("OrdType").Specific.DataBind.SetBound(true, "", "OrdType");

                //생산미완료(체크박스)
                oForm.DataSources.UserDataSources.Add("CmpltYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CmpltYN").Specific.DataBind.SetBound(true, "", "CmpltYN");

                //선택작번
                oForm.DataSources.UserDataSources.Add("SItemCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SItemCD").Specific.DataBind.SetBound(true, "", "SItemCD");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP750_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                //자체/외주
                oForm.Items.Item("InOut").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("IN", "자체");
                oForm.Items.Item("InOut").Specific.ValidValues.Add("OUT", "외주");
                oForm.Items.Item("InOut").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //장비/공구
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("M", "장비");
                oForm.Items.Item("ItemType").Specific.ValidValues.Add("T", "공구");
                oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //거래처구분
                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //조회구분
                oForm.Items.Item("SrchType").Specific.ValidValues.Add("1", "집계 보기");
                oForm.Items.Item("SrchType").Specific.ValidValues.Add("2", "상세 보기(구매)");
                oForm.Items.Item("SrchType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품의여부
                oForm.Items.Item("MM030YN").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("MM030YN").Specific.ValidValues.Add("Y", "완료");
                oForm.Items.Item("MM030YN").Specific.ValidValues.Add("N", "미완료");
                oForm.Items.Item("MM030YN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //입고여부
                oForm.Items.Item("MM050YN").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("MM050YN").Specific.ValidValues.Add("Y", "완료");
                oForm.Items.Item("MM050YN").Specific.ValidValues.Add("N", "미완료");
                oForm.Items.Item("MM050YN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //검수여부
                oForm.Items.Item("MM070YN").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("MM070YN").Specific.ValidValues.Add("Y", "완료");
                oForm.Items.Item("MM070YN").Specific.ValidValues.Add("N", "미완료");
                oForm.Items.Item("MM070YN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품의구분
                oForm.Items.Item("OrdType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdType").Specific, "SELECT Code, Name From [@PSH_ORDTYP] Order by Code", "", false, false);
                oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// Form 초기 세팅
        /// </summary>
        private void PS_PP750_InitializeForm()
        {
            oForm.DataSources.UserDataSources.Item("FrDocDt").Value = DateTime.Now.ToString("yyyy0101");
            oForm.DataSources.UserDataSources.Item("ToDocDt").Value = DateTime.Now.ToString("yyyy1231");
            oForm.DataSources.UserDataSources.Item("FrDueDt").Value = DateTime.Now.ToString("yyyy0101");
            oForm.DataSources.UserDataSources.Item("ToDueDt").Value = DateTime.Now.ToString("yyyy1231");
            oForm.Items.Item("CmpltYN").Width = 90;
            oForm.Items.Item("BtnPrint").Visible = false; //출력버튼 비활성
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PS_PP750_SelectData()
        {
            string sQry = string.Empty;
            string frDocDt; //수주일자(FR)
            string toDocDt; //수주일자(TO)
            string frDueDt; //납기일자(FR)
            string toDueDt; //납기일자(TO)
            string cntcCode; //생산담당
            string itemCode; //작번
            string inOut; //자체/외주
            string itemType; //장비/공구
            string cardType; //거래처구분
            string srchType; //조회구분(쿼리 매개변수 없음)
            string MM030YN; //품의여부
            string MM050YN; //입고여부
            string MM070YN; //검수여부
            string ordType; //품의구분
            string cmpltYN; //생산미완료여부
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                frDocDt = oForm.Items.Item("FrDocDt").Specific.Value.ToString().Trim(); //수주일자(FR)
                toDocDt = oForm.Items.Item("ToDocDt").Specific.Value.ToString().Trim(); //수주일자(TO)
                frDueDt = oForm.Items.Item("FrDueDt").Specific.Value.ToString().Trim(); //납기일자(FR)
                toDueDt = oForm.Items.Item("ToDueDt").Specific.Value.ToString().Trim(); //납기일자(TO)
                cntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim(); //생산담당
                itemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(); //작번
                inOut = oForm.Items.Item("InOut").Specific.Value.ToString().Trim(); //자체/외주
                itemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim(); //장비/공구
                cardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim(); //거래처구분
                srchType = oForm.Items.Item("SrchType").Specific.Value.ToString().Trim(); //조회구분(쿼리 매개변수 없음)
                MM030YN = oForm.Items.Item("MM030YN").Specific.Value.ToString().Trim(); //품의여부
                MM050YN = oForm.Items.Item("MM050YN").Specific.Value.ToString().Trim(); //입고여부
                MM070YN = oForm.Items.Item("MM070YN").Specific.Value.ToString().Trim(); //검수여부
                ordType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim(); //품의구분
                cmpltYN = (oForm.Items.Item("CmpltYN").Specific.Checked ? "Y" : "N"); //생산미완료여부

                oForm.DataSources.UserDataSources.Item("SItemCD").Value = ""; //선택작번 초기화

                //조회구분에 따라
                if (srchType == "1") //집계 보기
                {
                    sQry = "EXEC [PS_PP750_01] ";
                    sQry += "'" + frDocDt + "',";
                    sQry += "'" + toDocDt + "',";
                    sQry += "'" + frDueDt + "',";
                    sQry += "'" + toDueDt + "',";
                    sQry += "'" + cntcCode + "',";
                    sQry += "'" + itemCode + "',";
                    sQry += "'" + inOut + "',";
                    sQry += "'" + itemType + "',";
                    sQry += "'" + cardType + "',";
                    sQry += "'" + MM030YN + "',";
                    sQry += "'" + MM050YN + "',";
                    sQry += "'" + MM070YN + "',";
                    sQry += "'" + ordType + "',";
                    sQry += "'" + cmpltYN + "'";
                }
                else if(srchType == "2") //상세 보기(구매)
                {
                    sQry = "EXEC [PS_PP750_02] ";
                    sQry += "'" + frDocDt + "',";
                    sQry += "'" + toDocDt + "',";
                    sQry += "'" + frDueDt + "',";
                    sQry += "'" + toDueDt + "',";
                    sQry += "'" + cntcCode + "',";
                    sQry += "'" + itemCode + "',";
                    sQry += "'" + inOut + "',";
                    sQry += "'" + itemType + "',";
                    sQry += "'" + cardType + "',";
                    sQry += "'" + MM030YN + "',";
                    sQry += "'" + MM050YN + "',";
                    sQry += "'" + MM070YN + "',";
                    sQry += "'" + ordType + "',";
                    sQry += "'" + cmpltYN + "'";
                }

                mainGrid.DataTable.Clear();
                oDS_PS_PP750.ExecuteQuery(sQry);

                if (srchType == "1")
                {
                    mainGrid.Columns.Item(5).RightJustified = true; //수주수량(5)
                    mainGrid.Columns.Item(6).RightJustified = true; //수주금액(6)
                    mainGrid.Columns.Item(13).RightJustified = true; //작번등록횟수(13)
                    mainGrid.Columns.Item(16).RightJustified = true; //작업지시등록횟수(16)
                    mainGrid.Columns.Item(18).RightJustified = true; //작업일보횟수(18)
                    mainGrid.Columns.Item(20).RightJustified = true; //구매요청수량(20)
                    mainGrid.Columns.Item(21).RightJustified = true; //구매요청횟수(21)
                    mainGrid.Columns.Item(23).RightJustified = true; //구매견적수량(23)
                    mainGrid.Columns.Item(24).RightJustified = true; //구매견적횟수(24)
                    mainGrid.Columns.Item(26).RightJustified = true; //구매품의수량(26)
                    mainGrid.Columns.Item(27).RightJustified = true; //구매품의횟수(27)
                    mainGrid.Columns.Item(29).RightJustified = true; //가입고수량(29)
                    mainGrid.Columns.Item(30).RightJustified = true; //가입고횟수(30)
                    mainGrid.Columns.Item(32).RightJustified = true; //검수입고수량(32)
                    mainGrid.Columns.Item(33).RightJustified = true; //검수입고횟수(33)
                    mainGrid.Columns.Item(35).RightJustified = true; //검사횟수(35)
                    mainGrid.Columns.Item(37).RightJustified = true; //생산완료수량(37)
                    mainGrid.Columns.Item(38).RightJustified = true; //생산완료횟수(38)
                    mainGrid.Columns.Item(39).RightJustified = true; //생산잔량(수주-생산)(39)

                    //Grid 컬러 초기화
                    for (int i = 0; i < mainGrid.Columns.Count; i++)
                    {
                        mainGrid.Columns.Item(i).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(231, 231, 231));
                    }

                    //행 컬러 지정
                    for (int i = 0; i < mainGrid.Rows.Count; i++)
                    {
                        if (i % 2 == 1)
                        {
                            mainGrid.CommonSetting.SetRowBackColor(i + 1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(245, 245, 245)));
                        }
                    }

                    mainGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //납기일(9)
                    mainGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작번등록횟수(13)
                    mainGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업지시등록횟수(16)
                    mainGrid.Columns.Item(18).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업일보횟수(18)
                    mainGrid.Columns.Item(21).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매요청횟수(21)
                    mainGrid.Columns.Item(24).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매견적횟수(24)
                    mainGrid.Columns.Item(27).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //구매품의횟수(27)
                    mainGrid.Columns.Item(30).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //가입고횟수(30)
                    mainGrid.Columns.Item(33).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //검수입고횟수(33)
                    mainGrid.Columns.Item(35).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //검사횟수(35)
                    mainGrid.Columns.Item(38).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //생산완료횟수(38)

                    oForm.Items.Item("SItemCD").Enabled = true;
                    oForm.Items.Item("BtnPPDtl").Enabled = true;
                    oForm.Items.Item("BtnMMDtl").Enabled = true;
                }
                else if (srchType == "2")
                {
                    mainGrid.Columns.Item(5).RightJustified = true; //수주수량(5)
                    mainGrid.Columns.Item(6).RightJustified = true; //수주금액(6)
                    mainGrid.Columns.Item(13).RightJustified = true; //작번등록횟수(13)
                    mainGrid.Columns.Item(16).RightJustified = true; //작업지시등록횟수(16)
                    mainGrid.Columns.Item(18).RightJustified = true; //작업일보횟수(18)
                    mainGrid.Columns.Item(28).RightJustified = true; //청구수량(28)
                    mainGrid.Columns.Item(37).RightJustified = true; //품의금액(37)
                    mainGrid.Columns.Item(38).RightJustified = true; //품의수량(38)
                    mainGrid.Columns.Item(39).RightJustified = true; //견적수량(39)
                    mainGrid.Columns.Item(42).RightJustified = true; //입고수량(42)
                    mainGrid.Columns.Item(43).RightJustified = true; //검수수량(43)
                    mainGrid.Columns.Item(45).RightJustified = true; //검수금액(45)
                    mainGrid.Columns.Item(46).RightJustified = true; //미입고수량(46)
                    mainGrid.Columns.Item(47).RightJustified = true; //미입고금액(47)
                    mainGrid.Columns.Item(49).RightJustified = true; //검사횟수(49)
                    mainGrid.Columns.Item(51).RightJustified = true; //생산완료수량(51)
                    mainGrid.Columns.Item(52).RightJustified = true; //생산완료횟수(52)
                    mainGrid.Columns.Item(53).RightJustified = true; //생산잔량(수주-생산)(53)

                    //Grid 컬러 초기화
                    for (int i = 0; i < mainGrid.Columns.Count; i++)
                    {
                        mainGrid.Columns.Item(i).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(231, 231, 231));
                    }

                    //동일 작번 행 컬러 지정
                    for (int i = 0; i < mainGrid.Rows.Count; i++)
                    {
                        if (mainGrid.DataTable.GetValue(1, i) == "")
                        {
                            mainGrid.CommonSetting.SetRowBackColor(i + 1, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(245, 245, 245)));
                        }
                    }

                    mainGrid.Columns.Item(9).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //납기일(9)
                    mainGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작번등록횟수(13)
                    mainGrid.Columns.Item(16).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업지시등록횟수(16)
                    mainGrid.Columns.Item(18).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //작업일보횟수(18)
                    mainGrid.Columns.Item(35).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //거래처코드(35)
                    mainGrid.Columns.Item(36).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //거래처명(36)
                    mainGrid.Columns.Item(49).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //검사횟수(49)
                    mainGrid.Columns.Item(52).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204)); //생산완료횟수(52)

                    oForm.Items.Item("SItemCD").Enabled = false;
                    oForm.Items.Item("BtnPPDtl").Enabled = false;
                    oForm.Items.Item("BtnMMDtl").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// ResizeForm
        /// </summary>
        private void PS_PP750_ResizeForm()
        {
            try
            {
                if (mainGrid.Columns.Count > 0)
                {
                    mainGrid.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
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
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    // Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
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
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        PS_PP750_SelectData();
                    }
                    else if (pVal.ItemUID == "BtnPPDtl") //생산상세정보조회
                    {
                        PS_PP751 tempForm = new PS_PP751();
                        tempForm.LoadForm(this);
                    }
                    else if (pVal.ItemUID == "BtnMMDtl") //구매상세정보조회
                    {
                        PS_PP752 tempForm = new PS_PP752();
                        tempForm.LoadForm(this);
                    }
                    else if (pVal.ItemUID == "mainGrid") //그리드
                    {
                        if (mainGrid.Rows.SelectedRows.Count != 0)
                        {
                            if (oForm.DataSources.UserDataSources.Item("SrchType").Value == "1") //집계보기에서만 선택작번 연동
                            {
                                oForm.DataSources.UserDataSources.Item("SItemCD").Value = mainGrid.DataTable.GetValue(1, mainGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder));
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        if (pVal.ItemUID == "ItemCode" || pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
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
        /// ITEM_PRESSED 이벤트
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
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "ItemCode":
                                oForm.DataSources.UserDataSources.Item("ItemName").Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                oForm.DataSources.UserDataSources.Item("ItemSpec").Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                break;
                            case "CntcCode":
                                oForm.DataSources.UserDataSources.Item("CntcName").Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainGrid);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP750);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    PS_PP750_ResizeForm();
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
