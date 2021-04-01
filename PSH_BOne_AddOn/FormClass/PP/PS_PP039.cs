using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 1-나.작업지시-공정추가등록,수정,삭제
    /// </summary>
    internal class PS_PP039 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Grid oGrid01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP039L; //등록라인
        
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP039.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PS_PP039_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP039");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PP039_CreateItems();
                PS_PP039_ComboBox_Setting();
                PS_PP039_FormResize();
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
                oForm.Items.Item("OrdNum").Click();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP039_CreateItems()
        {
            try
            {
                oDS_PS_PP039L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //// 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oGrid01 = oForm.Items.Item("Grid01").Specific;

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

                //작업구분
                oForm.DataSources.UserDataSources.Add("OrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("OrdGbn").Specific.DataBind.SetBound(true, "", "OrdGbn");

                //작업지시일자(시작)
                oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");
                oForm.Items.Item("FrDt").Specific.VALUE = Convert.ToString(DateTime.Today.AddMonths(-2).ToString("yyyyMM01"));

                //작업지시일자(종료)
                oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
                oForm.Items.Item("ToDt").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

                //담당자
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //담당자명
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //작번
                oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
                oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

                //서브작번1
                oForm.DataSources.UserDataSources.Add("OrdSub1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
                oForm.Items.Item("OrdSub1").Specific.DataBind.SetBound(true, "", "OrdSub1");

                //서브작번2
                oForm.DataSources.UserDataSources.Add("OrdSub2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
                oForm.Items.Item("OrdSub2").Specific.DataBind.SetBound(true, "", "OrdSub2");

                //품명
                oForm.DataSources.UserDataSources.Add("JakMyung", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("JakMyung").Specific.DataBind.SetBound(true, "", "JakMyung");

                //규격
                oForm.DataSources.UserDataSources.Add("JakSize", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("JakSize").Specific.DataBind.SetBound(true, "", "JakSize");

                //그리드에서 선택한 행의 작업지시 문서번호
                oForm.DataSources.UserDataSources.Add("MainEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("MainEntry").Specific.DataBind.SetBound(true, "", "MainEntry");

                //그리드에서 선택한 행 번호
                oForm.DataSources.UserDataSources.Add("GridRow", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("GridRow").Specific.DataBind.SetBound(true, "", "GridRow");

                //그리드에서 선택한 작번(전체작번)
                oForm.DataSources.UserDataSources.Add("FullOrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("FullOrdNum").Specific.DataBind.SetBound(true, "", "FullOrdNum");

                //공정금액 합계
                oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_PRICE);
                oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");

                ////////////각 매트릭스 서식세팅 선택용 라디오버튼//////////S
                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                ////////////각 매트릭스 서식세팅 선택용 라디오버튼//////////E

                //참조정보 관련 컨트롤 숨김_S
                //Haeder
                oForm.Items.Item("Static90").Visible = false;
                oForm.Items.Item("Static91").Visible = false;
                oForm.Items.Item("Static92").Visible = false;
                oForm.Items.Item("MainEntry").Visible = false;
                oForm.Items.Item("GridRow").Visible = false;
                oForm.Items.Item("FullOrdNum").Visible = false;

                //Line
                oMat01.Columns.Item("VisOrder").Visible = false;
                oMat01.Columns.Item("Object").Visible = false;
                oMat01.Columns.Item("LogInst").Visible = false;
                oMat01.Columns.Item("U_LineNum").Visible = false;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP039_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                //사업장
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //작업구분
                sQry = "           SELECT     Code AS [Code], ";
                sQry = sQry + "               Name AS [Name]";
                sQry = sQry + " FROM      [@PSH_ITMBSORT]";
                sQry = sQry + " WHERE     U_PudYN = 'Y'";
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, sQry, "", false, false);
                oForm.Items.Item("OrdGbn").Specific.Select("105", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //기계공구 기본 선택

                //작업구분
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "10", "자가");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "20", "정밀");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkGbn", "30", "외주");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("WorkGbn"), "PS_PP039", "Mat01", "WorkGbn", false);

                //실적여부
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ResultYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ResultYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("ResultYN"), "PS_PP039", "Mat01", "ResultYN", false);

                //재작업여부
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReWorkYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReWorkYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("ReWorkYN"), "PS_PP039", "Mat01", "ReWorkYN", false);

                //일보여부
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReportYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "ReportYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("ReportYN"), "PS_PP039", "Mat01", "ReportYN", false);

                //작업여부
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP039", "Mat01", "WorkYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("WorkYN"), "PS_PP039", "Mat01", "WorkYN", false);
                ////////////매트릭스//////////
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP039_FormResize()
        {
            try
            {
                oForm.Items.Item("Grid01").Height = oForm.Height / 2 - 100;
                oForm.Items.Item("Grid01").Width = oForm.Width - 30;

                if (oGrid01.Columns.Count > 0)
                {
                    oGrid01.AutoResizeColumns();
                }

                oForm.Items.Item("Opt02").Top = oForm.Height / 2 + 10;

                oForm.Items.Item("Static93").Top = oForm.Height / 2 + 10;
                oForm.Items.Item("Static93").Left = oForm.Items.Item("Opt02").Width + 955;
                oForm.Items.Item("Total").Top = oForm.Items.Item("Static93").Top;
                oForm.Items.Item("Total").Left = oForm.Items.Item("Static93").Left + oForm.Items.Item("Static93").Width;

                oForm.Items.Item("Static94").Left = oForm.Items.Item("Opt02").Width + 50;
                oForm.Items.Item("Static94").Top = oForm.Items.Item("Static93").Top;
                oForm.Items.Item("Static94").Left = oForm.Items.Item("Opt02").Width + 50;

                oForm.Items.Item("Mat01").Top = oForm.Items.Item("Opt02").Top + 15;
                oForm.Items.Item("Mat01").Height = oForm.Items.Item("Mat01").Top - 120;
                oForm.Items.Item("Mat01").Width = oForm.Width - 30;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP039_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP039_AddMatrixRow(int oRow, bool RowIserted = false)
        {
            try
            {
                oForm.Freeze(true);
                ////행추가여부
                if (RowIserted == false)
                {
                    oDS_PS_PP039L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP039L.Offset = oRow;
                oDS_PS_PP039L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP039L.SetValue("U_ColReg02", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// PS_PP038_MTX01
        /// </summary>
        private void PS_PP039_MTX01()
        {
            string errMessage = string.Empty;
            string BPLID;        //사업장
            string OrdGbn;       //작업구분
            string FrDt;         //지시일자(Fr)
            string ToDt;         //지시일자(To)
            string CntcCode;     //담당자
            string OrdNum;       //작번
            string OrdSub1;      //서브작번1
            string OrdSub2;      //서브작번2
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                ProgressBar01.Text = "조회중!";
                BPLID = oForm.Items.Item("BPLID").Specific.VALUE.ToString().Trim();                //사업장
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.VALUE.ToString().Trim();                //작업구분
                FrDt = oForm.Items.Item("FrDt").Specific.VALUE.ToString().Trim();                //지시일자(Fr)
                ToDt = oForm.Items.Item("ToDt").Specific.VALUE.ToString().Trim();                //지시일자(To)
                CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE.ToString().Trim();                //담당자
                OrdNum = oForm.Items.Item("OrdNum").Specific.VALUE.ToString().Trim();                //작번
                OrdSub1 = oForm.Items.Item("OrdSub1").Specific.VALUE.ToString().Trim();                //서브작번1
                OrdSub2 = oForm.Items.Item("OrdSub2").Specific.VALUE.ToString().Trim();                //서브작번2

                Query01 = "         EXEC PS_PP039_01 '";
                Query01 = Query01 + BPLID + "','";     //사업장
                Query01 = Query01 + OrdGbn + "','";    //작업구분
                Query01 = Query01 + FrDt + "','";      //지시일자(Fr)
                Query01 = Query01 + ToDt + "','";      //지시일자(To)
                Query01 = Query01 + CntcCode + "','";  //담당자
                Query01 = Query01 + OrdNum + "','";    //작번
                Query01 = Query01 + OrdSub1 + "','";   //서브작번1
                Query01 = Query01 + OrdSub2 + "'";     //서브작번2

                oGrid01.DataTable.Clear();
                oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01);
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");

                if (oGrid01.Rows.Count == 0)
                {
                    errMessage = "결과가 존재하지 않습니다";
                    throw new Exception();
                }
                oGrid01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (errMessage != null)
                {
                    PSH_Globals.SBO_Application.MessageBox("errMessage");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP039_MTX02
        /// </summary>
        private void PS_PP039_MTX02(int pRow)
        {
            short i;
            double TotalAmt = 0;
            string sQry;
            string errMessage = string.Empty;
            string DocEntry;
            string FullOrdNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                ProgressBar01.Text = "조회중!";
                DocEntry = oGrid01.DataTable.Columns.Item("문서번호").Cells.Item(pRow).Value.ToString().Trim();                //그리드에서 선택한 작업지시등록 문서번호
                FullOrdNum = oGrid01.DataTable.Columns.Item("작번").Cells.Item(pRow).Value.ToString().Trim() + "-" + oGrid01.DataTable.Columns.Item("서브작번1").Cells.Item(pRow).Value.ToString().Trim() + "-" + oGrid01.DataTable.Columns.Item("서브작번2").Cells.Item(pRow).Value.ToString().Trim();
                //그리드에서 선택한 작번(전체작번)

                oForm.Items.Item("MainEntry").Specific.VALUE = DocEntry;          //그리드에서 선택한 행의 작업지시 문서번호 레이블에 저장
                oForm.Items.Item("GridRow").Specific.VALUE = pRow;                //그리드에서 선택한 행의 행번호
                oForm.Items.Item("FullOrdNum").Specific.VALUE = FullOrdNum;       //그리드에서 선택한 작번(전체작번)

                sQry = "         EXEC [PS_PP039_02] '";
                sQry = sQry + DocEntry + "'";

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_PP039L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP039L.Size)
                    {
                        oDS_PS_PP039L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_PP039L.Offset = i;

                    oDS_PS_PP039L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP039L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim());                    //선택
                    oDS_PS_PP039L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("Sequence").Value.ToString().Trim());                    //순번
                    oDS_PS_PP039L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("CpBCode").Value.ToString().Trim());                    //공정대분류
                    oDS_PS_PP039L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("CpBName").Value.ToString().Trim());                    //대분류명
                    oDS_PS_PP039L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("CpCode").Value.ToString().Trim());                    //공정중분류
                    oDS_PS_PP039L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("CpName").Value.ToString().Trim());                    //중분류명
                    oDS_PS_PP039L.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("StdHour").Value.ToString().Trim());                    //표준공수
                    oDS_PS_PP039L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("Unit").Value.ToString().Trim());                    //단위
                    oDS_PS_PP039L.SetValue("U_ColPrc01", i, oRecordSet01.Fields.Item("CpPrice").Value.ToString().Trim());                    //공정금액
                    oDS_PS_PP039L.SetValue("U_ColDt01", i,  oRecordSet01.Fields.Item("ReDate").Value.ToString().Trim());                    //완료요구일
                    oDS_PS_PP039L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("WorkGbn").Value.ToString().Trim());                    //작업구분
                    oDS_PS_PP039L.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("ReWorkYN").Value.ToString().Trim());                    //재작업여부
                    oDS_PS_PP039L.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("ResultYN").Value.ToString().Trim());                    //실적여부
                    oDS_PS_PP039L.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("ReportYN").Value.ToString().Trim());                    //일보여부
                    oDS_PS_PP039L.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("FailCode").Value.ToString().Trim());                    //재작업사유
                    oDS_PS_PP039L.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("FailName").Value.ToString().Trim());                    //재작업사유명
                    oDS_PS_PP039L.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("WorkYN").Value.ToString().Trim());                    //작업여부
                    oDS_PS_PP039L.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());                    //DocEntry
                    oDS_PS_PP039L.SetValue("U_ColNum02", i, oRecordSet01.Fields.Item("LineId").Value.ToString().Trim());                    //LineId
                    oDS_PS_PP039L.SetValue("U_ColNum03", i, oRecordSet01.Fields.Item("VisOrder").Value.ToString().Trim());                    //VisOrder
                    oDS_PS_PP039L.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("Object").Value.ToString().Trim());                    //Object
                    oDS_PS_PP039L.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("LogInst").Value.ToString().Trim());                    //LogInst
                    oDS_PS_PP039L.SetValue("U_ColNum04", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());                    //LineNum

                    TotalAmt = TotalAmt + Convert.ToDouble(oRecordSet01.Fields.Item("CpPrice").Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                //합계 계산_S
                oForm.Items.Item("Total").Specific.VALUE = TotalAmt;
                //합계 계산_E
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (errMessage != null)
                {
                    PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP039_CheckBeforeSearch
        /// </summary>
        private bool PS_PP039_CheckBeforeSearch()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;
            try
            {
                if (oForm.Items.Item("OrdGbn").Specific.VALUE.ToString().Trim() == "%")
                {
                    errMessage = "조회조건 작업구분은 필수선택 사항입니다. 확인하세요.";
                    oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != null)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP039_CheckOKYN
        /// </summary>
        private bool PS_PP039_CheckOKYN(int pRow)
        {
            bool functionReturnValue = false;
            string PP030DL;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                PP030DL = oDS_PS_PP039L.GetValue("U_ColNum01", pRow - 1) + "-" + oDS_PS_PP039L.GetValue("U_ColNum02", pRow - 1);

                Query01 = "           SELECT    U_OKYN AS [OKYN]";
                Query01 = Query01 + " FROM      [@PS_MM005H] ";
                Query01 = Query01 + " WHERE     U_PP030DL = '" + PP030DL + "'";
                Query01 = Query01 + "           AND U_OrdType = '10'";
                //원자재 구매요청만 조회

                oRecordSet01.DoQuery(Query01);

                if (oRecordSet01.Fields.Item("OKYN").Value == "Y")
                {
                    functionReturnValue = true;
                }
                else
                {
                    functionReturnValue = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP038_CheckBeforeSearch
        /// </summary>
        private bool PS_PP039_Check_DupReq(string pDocEntry, string pItemCode, string pLineID)
        {
            bool functionReturnValue = false;
            string Query01;
            string DocEntry;
            string ItemCode;
            string LineId;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                DocEntry = pDocEntry;
                ItemCode = pItemCode;
                LineId = pLineID;

                Query01 = "         EXEC PS_Z_Check_DupReq '";
                Query01 = Query01 + DocEntry + "','";
                Query01 = Query01 + ItemCode + "','";
                Query01 = Query01 + LineId + "'";

                oRecordSet01.DoQuery(Query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "false")
                {
                    functionReturnValue = false;
                }
                else
                {
                    functionReturnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private bool PS_PP039_AddData()
        {
            bool functionReturnValue = false;
            short loopCount;
            string sQry;
            string Sequence;       //순번
            string CpBCode;        //공정대분류
            string CpBName;        //대분류명
            string CpCode;         //공정중분류
            string CpName;         //중분류명
            string StdHour;        //표준공수
            string Unit;           //단위
            double CpPrice;        //공정금액
            string ReDate;         //완료요구일
            string WorkGbn;        //작업구분
            string ReWorkYN;       //재작업여부
            string ResultYN;       //실적여부
            string ReportYN ;      //일보여부
            string FailCode;       //재작업사유
            string FailName;       //재작업사유명
            string WorkYN;         //작업여부
            string DocEntry;       //DocEntry
            string LineId;         //LineID
            string VisOrder;       //VisOrder
            string Object_Renamed; //Object
            string LogInst;        //LogInst
            string LineNum;        //U_LineNum
            string MainEntry;      //작업지시문서번호
            string BPLID;          //사업장코드
            string FullOrdNum;      //전체작번
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                MainEntry = oForm.Items.Item("MainEntry").Specific.VALUE.ToString().Trim();
                BPLID = oForm.Items.Item("BPLID").Specific.VALUE.ToString().Trim();
                FullOrdNum = oForm.Items.Item("FullOrdNum").Specific.VALUE.ToString().Trim();

                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_PP039L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        Sequence = oDS_PS_PP039L.GetValue("U_ColReg02", loopCount).ToString().Trim();                  //순번
                        CpBCode =  oDS_PS_PP039L.GetValue("U_ColReg03", loopCount).ToString().Trim();                  //공정대분류
                        CpBName =  oDS_PS_PP039L.GetValue("U_ColReg04", loopCount).ToString().Trim();                  //대분류명
                        CpCode =   oDS_PS_PP039L.GetValue("U_ColReg05", loopCount).ToString().Trim();                  //공정중분류
                        CpName =   oDS_PS_PP039L.GetValue("U_ColReg06", loopCount).ToString().Trim();                  //중분류명
                        StdHour =  oDS_PS_PP039L.GetValue("U_ColQty01", loopCount).ToString().Trim();                  //표준공수
                        Unit =     oDS_PS_PP039L.GetValue("U_ColReg08", loopCount).ToString().Trim();                  //단위
                        CpPrice = Convert.ToDouble(oDS_PS_PP039L.GetValue("U_ColPrc01", loopCount).ToString().Trim()); //공정금액
                        ReDate =   oDS_PS_PP039L.GetValue("U_ColDt01", loopCount).ToString().Trim();                   //완료요구일
                        WorkGbn =  oDS_PS_PP039L.GetValue("U_ColReg11", loopCount).ToString().Trim();                  //작업구분
                        ReWorkYN = oDS_PS_PP039L.GetValue("U_ColReg12", loopCount).ToString().Trim();                  //재작업여부
                        ResultYN = oDS_PS_PP039L.GetValue("U_ColReg13", loopCount).ToString().Trim();                  //실적여부
                        ReportYN = oDS_PS_PP039L.GetValue("U_ColReg14", loopCount).ToString().Trim();                  //일보여부
                        FailCode = oDS_PS_PP039L.GetValue("U_ColReg15", loopCount).ToString().Trim();                  //재작업사유
                        FailName = oDS_PS_PP039L.GetValue("U_ColReg16", loopCount).ToString().Trim();                  //재작업사유명
                        WorkYN =  oDS_PS_PP039L.GetValue("U_ColReg17", loopCount).ToString().Trim();                   //작업여부
                        DocEntry = oDS_PS_PP039L.GetValue("U_ColNum01", loopCount).ToString().Trim();                  //DocEntry
                        LineId = oDS_PS_PP039L.GetValue("U_ColNum02", loopCount).ToString().Trim();                    //LineId
                        VisOrder = oDS_PS_PP039L.GetValue("U_ColNum03", loopCount).ToString().Trim();                  //VisOrder
                        Object_Renamed = oDS_PS_PP039L.GetValue("U_ColReg18", loopCount).ToString().Trim();            //Object
                        LogInst = oDS_PS_PP039L.GetValue("U_ColReg19", loopCount).ToString().Trim();                   //LogInst
                        LineNum = oDS_PS_PP039L.GetValue("U_ColNum04", loopCount).ToString().Trim();                   //U_LineNum

                        sQry = "                EXEC [PS_PP039_03] ";
                        sQry = sQry + "'" + Sequence + "',";      //순번
                        sQry = sQry + "'" + CpBCode + "',";       //공정대분류
                        sQry = sQry + "'" + CpBName + "',";       //대분류명
                        sQry = sQry + "'" + CpCode + "',";        //공정중분류
                        sQry = sQry + "'" + CpName + "',";        //중분류명
                        sQry = sQry + "'" + StdHour + "',";       //표준공수
                        sQry = sQry + "'" + Unit + "',";          //단위
                        sQry = sQry + "'" + CpPrice + "',";       //공정금액
                        sQry = sQry + "'" + ReDate + "',";        //완료요구일
                        sQry = sQry + "'" + WorkGbn + "',";       //작업구분
                        sQry = sQry + "'" + ReWorkYN + "',";      //재작업여부
                        sQry = sQry + "'" + ResultYN + "',";      //실적여부
                        sQry = sQry + "'" + ReportYN + "',";      //일보여부
                        sQry = sQry + "'" + FailCode + "',";      //재작업사유
                        sQry = sQry + "'" + FailName + "',";      //재작업사유명
                        sQry = sQry + "'" + WorkYN + "',";        //작업여부
                        sQry = sQry + "'" + DocEntry + "',";      //DocEntry
                        sQry = sQry + "'" + LineId + "',";        //LineID
                        sQry = sQry + "'" + VisOrder + "',";      //VisOrder
                        sQry = sQry + "'" + Object_Renamed + "',";//Object
                        sQry = sQry + "'" + LogInst + "',";       //LogInst
                        sQry = sQry + "'" + LineNum + "',";       //LineNum
                        sQry = sQry + "'" + MainEntry + "',";     //선택한 작업지시문서번호
                        sQry = sQry + "'" + BPLID + "',";         //사업장코드
                        sQry = sQry + "'" + FullOrdNum + "'";     //전체작번

                        oRecordSet01.DoQuery(sQry);
                    }
                }
                PSH_Globals.SBO_Application.MessageBox("등록완료!");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private bool PS_PP039_DeleteData()
        {
            bool functionReturnValue = false;
            short loopCount;
            string sQry;
            string DocEntry;//문서번호
            string LineId;  //라인번호
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "삭제대상이 없습니다. 확인하세요.";
                    throw new Exception();
                }
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    //체크된 행만 'And Trim(oDS_PS_PP039L.GetValue("U_ColReg17", loopCount)) = "N" Then '선택된 행 중에 작업일보가 등록되지 않은 공정만 삭제
                    if (oDS_PS_PP039L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        DocEntry = oDS_PS_PP039L.GetValue("U_ColNum01", loopCount).ToString().Trim();//문서번호
                        LineId = oDS_PS_PP039L.GetValue("U_ColNum02", loopCount).ToString().Trim();  //라인번호

                        sQry = "EXEC [PS_PP039_05] ";        //외주제작청구되었는지 체크
                        sQry = sQry + "'" + DocEntry + "',"; //문서번호
                        sQry = sQry + "'" + LineId + "'";    //라인번호
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item("CNT").Value > 0)
                        {
                            errMessage = loopCount + 1 + "행(" + DocEntry + "-" + LineId + ")은 외주제작청구가 등록되었습니다. 삭제할 수 없습니다. 삭제명령은 중단됩니다.";
                            throw new Exception();
                        }

                        sQry = "EXEC [PS_PP039_06] ";          //작업일보가 등록되었는지 체크
                        sQry = sQry + "'" + DocEntry + "',";   //문서번호
                        sQry = sQry + "'" + LineId + "'";      //라인번호
                        oRecordSet01.DoQuery(sQry);

                        if (oRecordSet01.Fields.Item("CNT").Value > 0)
                        {
                            errMessage = loopCount + 1 + "행(" + DocEntry + "-" + LineId + ")은 작업일보가 등록되었습니다. 삭제할 수 없습니다. 삭제명령은 중단됩니다.";
                            throw new Exception();
                        }

                        sQry = "EXEC [PS_PP039_04] ";
                        sQry = sQry + "'" + DocEntry + "',";                        //문서번호
                        sQry = sQry + "'" + LineId + "'";                        //라인번호
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != null)
                {
                    PSH_Globals.SBO_Application.MessageBox("errMessage");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP039_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP039'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP039_DataValidCheck()
        {
            bool functionReturnValue = false;
            int loopCount;
            string CpBCode;  //공정대분류
            string CpCode;   //공정중분류              
            string WorkGbn;  //작업구분
            string ReWorkYN; //재작업여부
            string ResultYN; //실적여부
            string ReportYN; //일보여부
            string errMessage = string.Empty;
            try
            {
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_PP039L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        CpBCode = oDS_PS_PP039L.GetValue("U_ColReg03", loopCount).ToString().Trim();  //공정대분류
                        CpCode = oDS_PS_PP039L.GetValue("U_ColReg05", loopCount).ToString().Trim();   //공정중분류
                        WorkGbn = oDS_PS_PP039L.GetValue("U_ColReg11", loopCount).ToString().Trim();  //작업구분
                        ReWorkYN = oDS_PS_PP039L.GetValue("U_ColReg12", loopCount).ToString().Trim(); //재작업여부
                        ResultYN = oDS_PS_PP039L.GetValue("U_ColReg13", loopCount).ToString().Trim(); //실적여부
                        ReportYN = oDS_PS_PP039L.GetValue("U_ColReg14", loopCount).ToString().Trim(); //일보여부

                        if (string.IsNullOrEmpty(CpBCode))
                        {
                            errMessage = "공정대분류는 필수입니다.";
                            oMat01.Columns.Item("CpBCode").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(CpCode))
                        {
                            errMessage = "공정중분류는 필수입니다.";
                            oMat01.Columns.Item("CpCode").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(WorkGbn))
                        {
                            errMessage = "작업구분은 필수입니다.";
                            oMat01.Columns.Item("WorkGbn").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(ReWorkYN))
                        {
                            errMessage = "재작업여부는 필수입니다.";
                            oMat01.Columns.Item("ReWorkYN").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(ResultYN))
                        {
                            errMessage = "실적여부는 필수입니다.";
                            oMat01.Columns.Item("ResultYN").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(ReportYN))
                        {
                            errMessage = "일보여부는 필수입니다.";
                            oMat01.Columns.Item("ReportYN").Cells.Item(loopCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != null)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP039_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
        {
            short loopCount;
            string sQry ;
            string OrdNum ;
            string OrdSub1;
            string OrdSub2;
            double TotalAmt = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (oUID == "Mat01")
                {
                    if (oCol == "CpBCode")
                    {
                        oMat01.FlushToDataSource();
                        oDS_PS_PP039L.SetValue("U_ColReg04", oRow - 1, dataHelpClass.GetValue("SELECT Name FROM [@PS_PP001H] WHERE Code = '" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE + "'", 0, 1));
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == oRow & !string.IsNullOrEmpty(oDS_PS_PP039L.GetValue("U_ColReg02", oRow - 1).ToString().Trim()))
                        {
                            PS_PP039_AddMatrixRow(oRow);
                        }
                        oMat01.Columns.Item("CpBCode").Cells.Item(oRow).Click();
                        //공정대분류 입력 후 TAB 조정
                    }
                    else if (oCol == "CpCode")
                    {
                        oMat01.FlushToDataSource();
                        oDS_PS_PP039L.SetValue("U_ColReg06", oRow - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE Code = '" + oMat01.Columns.Item("CpBCode").Cells.Item(oRow).Specific.VALUE + "' AND U_CpCode = '" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE + "'", 0, 1));
                        //작업구분은 기본으로 "자가" 선택
                        oDS_PS_PP039L.SetValue("U_ColReg11", oRow - 1, "10");

                        //공정선택 시 재작업여부 자동 선택
                        //PK/탈지일때 재작업여부 "예"
                        if (oMat01.Columns.Item("CpCode").Cells.Item(oRow).Specific.VALUE == "CP50103" | oMat01.Columns.Item("CpCode").Cells.Item(oRow).Specific.VALUE == "CP50106")
                        {
                            oDS_PS_PP039L.SetValue("U_ColReg12", oRow - 1, "Y");
                        }
                        else
                        {
                            oDS_PS_PP039L.SetValue("U_ColReg12", oRow - 1, "N");
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.Columns.Item("CpCode").Cells.Item(oRow).Click();
                        //공정중분류 입력 후 TAB 조정
                    }
                    else if (oCol == "FailCode")
                    {
                        oMat01.FlushToDataSource();
                        oDS_PS_PP039L.SetValue("U_ColReg16", oRow - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat01.Columns.Item("FailCode").Cells.Item(oRow).Specific.VALUE + "'", 0, 1));
                        oMat01.LoadFromDataSource();
                        oMat01.Columns.Item("FailCode").Cells.Item(oRow).Click();
                        //재작업사유 입력 후 TAB 조정
                    }

                    if (oCol == "StdHour" | oCol == "ReDate")
                    {
                        oMat01.FlushToDataSource();
                        //표준공수와 완료요구일은 수정이 가능해야 하므로 Flush 를 함
                        //표준공수 등록 시
                        if (oCol == "StdHour")
                        {
                            //공정단가 계산_S
                            if (oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.VALUE == "10")
                            {
                                oDS_PS_PP039L.SetValue("U_ColPrc01", oRow - 1, Convert.ToString(Convert.ToDouble(dataHelpClass.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(oRow).Specific.VALUE + "'", 0, 1)) * Convert.ToDouble(oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE)));
                            }
                            else if (oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.VALUE == "20")
                            {
                                oDS_PS_PP039L.SetValue("U_ColPrc01", oRow - 1, Convert.ToString(dataHelpClass.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(oRow).Specific.VALUE + "'", 0, 1) * oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE));
                            }
                            oDS_PS_PP039L.SetValue("U_ColQty01", oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE);
                            //공정단가 계산_E

                            //합계 계산_S
                            for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                            {
                                TotalAmt += Convert.ToDouble(oDS_PS_PP039L.GetValue("U_ColPrc01", loopCount));
                            }
                            oForm.Items.Item("Total").Specific.VALUE = TotalAmt;
                            //합계 계산_E
                        }
                        oMat01.LoadFromDataSource();

                        if (oCol == "StdHour")
                        {
                            oMat01.Columns.Item("StdHour").Cells.Item(oRow).Click(); //표준공수 입력 후 TAB 조정
                        }
                        else if (oCol == "ReDate")
                        {
                            oMat01.Columns.Item("ReDate").Cells.Item(oRow).Click();  //완료요구일 입력 후 TAB 조정
                        }
                    }
                    oMat01.AutoResizeColumns();
                }
                else if (oUID == "CntcCode")
                {
                    oForm.Items.Item("CntcName").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.VALUE + "'","");
                }
                else if (oUID == "OrdNum" | oUID == "OrdSub1" | oUID == "OrdSub2")
                {
                    OrdNum =  oForm.Items.Item("OrdNum").Specific.VALUE.ToString().Trim();
                    OrdSub1 = oForm.Items.Item("OrdSub1").Specific.VALUE.ToString().Trim();
                    OrdSub2 = oForm.Items.Item("OrdSub2").Specific.VALUE.ToString().Trim();
                    
                    sQry = "           SELECT   CASE";
                    sQry = sQry + "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
                    sQry = sQry + "                 ELSE T0.U_JakMyung";
                    sQry = sQry + "             END AS [JakMyung],";
                    sQry = sQry + "             CASE";
                    sQry = sQry + "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
                    sQry = sQry + "                 ELSE T0.U_JakSize";
                    sQry = sQry + "             END AS [JakSize]";
                    sQry = sQry + " FROM     [@PS_PP020H] AS T0";
                    sQry = sQry + " WHERE   T0.U_JakName = '" + OrdNum + "'";
                    sQry = sQry + "             AND T0.U_SubNo1 = CASE WHEN '" + OrdSub1 + "' = '' THEN '00' ELSE '" + OrdSub1 + "' END";
                    sQry = sQry + "             AND T0.U_SubNo2 = CASE WHEN '" + OrdSub2 + "' = '' THEN '000' ELSE '" + OrdSub2 + "' END";

                    oRecordSet01.DoQuery(sQry);

                    oForm.Items.Item("JakMyung").Specific.VALUE = oRecordSet01.Fields.Item("JakMyung").Value;
                    oForm.Items.Item("JakSize").Specific.VALUE = oRecordSet01.Fields.Item("JakSize").Value;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                            if (PS_PP039_CheckBeforeSearch() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP039_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnAdd")
                    {
                        //필수 입력 사항 체크
                        if (PS_PP039_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            if (PS_PP039_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP039_MTX02(Convert.ToInt32(oForm.Items.Item("GridRow").Specific.VALUE));
                            PS_PP039_AddMatrixRow(oMat01.RowCount);
                        }
                    }
                    else if (pVal.ItemUID == "BtnDel")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            if (PS_PP039_DeleteData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP039_MTX02(oForm.Items.Item("GridRow").Specific.VALUE);
                            PS_PP039_AddMatrixRow(oMat01.RowCount);
                        }
                        else
                        {
                        }
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
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "CpBCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CpBCode");
                        }
                        else if (pVal.ColUID == "CpCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CpCode");
                        }
                        else if (pVal.ColUID == "FailCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "FailCode");
                        }
                    }
                    else
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", ""); //담당자
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");   //작번
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Grid01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oForm.Freeze(false);
                    }
                    if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oForm.Freeze(false);
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            PS_PP039_MTX02(pVal.Row);
                            PS_PP039_AddMatrixRow(oMat01.RowCount);
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
                    if (pVal.ItemChanged == true)
                    {
                        if ((pVal.ItemUID == "Mat01"))
                        {
                            PS_PP039_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                        else
                        {
                            PS_PP039_FlushToItemValue(pVal.ItemUID);
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
                    PS_PP039_AddMatrixRow(oMat01.VisualRowCount);
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
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
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
                    PS_PP039_FormResize();
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

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
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
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
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
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
                if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
