using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 1-가.작업지시-투입자재추가등록,수정,삭제
    /// </summary>
    internal class PS_PP038 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Grid oGrid01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP038L; //등록라인
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP038.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP038_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP038");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PP038_CreateItems();
                PS_PP038_ComboBox_Setting();
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
        private void PS_PP038_CreateItems()
        {
            try
            {
                oDS_PS_PP038L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
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
                oForm.Items.Item("FrDt").Specific.Value = Convert.ToString(DateTime.Today.AddMonths(-2).ToString("yyyyMM01"));

                //작업지시일자(종료)
                oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
                oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

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

                //각 매트릭스 서식세팅 선택용 라디오버튼//S
                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                //각 매트릭스 서식세팅 선택용 라디오버튼//E

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
                //참조정보 관련 컨트롤 숨김_E
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP038_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //작업구분
                sQry = "        SELECT    Code AS [Code], ";
                sQry += "           Name AS [Name]";
                sQry += " FROM      [@PSH_ITMBSORT]";
                sQry += " WHERE     U_PudYN = 'Y'";

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, sQry, "", false, false);
                oForm.Items.Item("OrdGbn").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //투입구분
                oMat01.Columns.Item("InputGbn").ValidValues.Add("10", "일반");
                oMat01.Columns.Item("InputGbn").ValidValues.Add("20", "원재료");
                oMat01.Columns.Item("InputGbn").ValidValues.Add("30", "스크랩");

                //품목그룹
                sQry = "        SELECT  ItmsGrpCod AS [Code], ";
                sQry += "         ItmsGrpNam AS [Name]";
                sQry += " FROM    [OITB] a";

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemGpCd"), sQry,"","");

                //조달방식
                oMat01.Columns.Item("ProcType").ValidValues.Add("10", "청구");
                oMat01.Columns.Item("ProcType").ValidValues.Add("20", "잔재");
                oMat01.Columns.Item("ProcType").ValidValues.Add("30", "취소");

                //수입품여부
                oMat01.Columns.Item("ImportYN").ValidValues.Add("Y", "Y");
                oMat01.Columns.Item("ImportYN").ValidValues.Add("N", "N");

                //긴급여부
                oMat01.Columns.Item("EmergYN").ValidValues.Add("Y", "Y");
                oMat01.Columns.Item("EmergYN").ValidValues.Add("N", "N");

                //청구사유(라인)
                sQry = "        SELECT      U_Minor,";
                sQry += "             U_CdName";
                sQry += " FROM        [@PS_SY001L]";
                sQry += " WHERE       Code = 'P203'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += "             AND U_Minor <> 'A'";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RCode"), sQry,"","");
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
        /// FormResize
        /// </summary>
        private void PS_PP038_FormResize()
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
        /// PS_PP038_CheckBeforeSearch
        /// </summary>
        private bool PS_PP038_CheckBeforeSearch()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if(oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "%")
                {
                    errMessage = "조회조건 작업구분은 필수선택 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
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
        /// PS_PP038_CheckOKYN
        /// </summary>
        /// <param name="pRow"></param>
        /// <returns></returns>
        private bool PS_PP038_CheckOKYN(int pRow)
        {
            bool functionReturnValue = false;
            string PP030DL;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                PP030DL = oDS_PS_PP038L.GetValue("U_ColNum01", pRow - 1) + "-" + oDS_PS_PP038L.GetValue("U_ColNum02", pRow - 1);

                Query01 = "           SELECT    U_OKYN AS [OKYN]";
                Query01 += " FROM      [@PS_MM005H] ";
                Query01 += " WHERE     U_PP030DL = '" + PP030DL + "'";
                Query01 += "           AND U_OrdType = '10'";

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
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP038_Check_DupReq
        /// </summary>
        private bool PS_PP038_Check_DupReq(string pDocEntry, string pItemCode, string pLineID)
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
                Query01 += DocEntry + "','";
                Query01 += ItemCode + "','";
                Query01 += LineId + "'";

                oRecordSet01.DoQuery(Query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "FALSE")
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
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP038_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP038_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                //행추가여부
                if (RowIserted == false)
                {
                    oDS_PS_PP038L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP038L.Offset = oRow;
                oDS_PS_PP038L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP038L.SetValue("U_ColReg02", oRow, "10");
                //투입구분
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
        private void PS_PP038_MTX01()
        {
            string BPLId;        //사업장
            string OrdGbn;       //작업구분
            string FrDt;         //지시일자(Fr)
            string ToDt;         //지시일자(To)
            string CntcCode;     //담당자
            string OrdNum;       //작번
            string OrdSub1;      //서브작번1
            string OrdSub2;      //서브작번2
            string Query01;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgressBar01.Text = "조회중!";
                BPLId =   oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();     //사업장
                OrdGbn =  oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();    //작업구분
                FrDt =    oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();      //지시일자(Fr)
                ToDt =    oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();      //지시일자(To)
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim(); //담당자
                OrdNum =  oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();    //작번
                OrdSub1 = oForm.Items.Item("OrdSub1").Specific.Value.ToString().Trim();   //서브작번1
                OrdSub2 = oForm.Items.Item("OrdSub2").Specific.Value.ToString().Trim();   //서브작번2

                Query01 = "         EXEC PS_PP038_01 '";
                Query01 += BPLId + "','";   //사업장
                Query01 += OrdGbn + "','";  //작업구분
                Query01 += FrDt + "','";    //지시일자(Fr)
                Query01 += ToDt + "','";    //지시일자(To)
                Query01 += CntcCode + "','";//담당자
                Query01 += OrdNum + "','";  //작번
                Query01 += OrdSub1 + "','"; //서브작번1
                Query01 += OrdSub2 + "'";   //서브작번2

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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP038_MTX02
        /// </summary>
        private void PS_PP038_MTX02(int pRow)
        {
            short i;
            string sQry;
            string DocEntry;
            string FullOrdNum;
            string errMessage = string.Empty; 
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgressBar01.Text = "조회중!";
                DocEntry = oGrid01.DataTable.Columns.Item("문서번호").Cells.Item(pRow).Value.ToString().Trim();   //그리드에서 선택한 작업지시등록 문서번호
                FullOrdNum = oGrid01.DataTable.Columns.Item("작번").Cells.Item(pRow).Value.ToString().Trim() + "-" +oGrid01.DataTable.Columns.Item("서브작번1").Cells.Item(pRow).Value.ToString().Trim() + "-" + oGrid01.DataTable.Columns.Item("서브작번2").Cells.Item(pRow).Value.ToString().Trim();

                oForm.Items.Item("MainEntry").Specific.Value = DocEntry;   //그리드에서 선택한 행의 작업지시 문서번호 레이블에 저장
                oForm.Items.Item("GridRow").Specific.Value = pRow;         //그리드에서 선택한 행의 행번호
                oForm.Items.Item("FullOrdNum").Specific.Value = FullOrdNum;//그리드에서 선택한 작번(전체작번)

                sQry = "      EXEC [PS_PP038_02] '";
                sQry += DocEntry + "'";

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_PP038L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP038L.Size)
                    {
                        oDS_PS_PP038L.InsertRecord(i);
                    }
                    oMat01.AddRow();
                    oDS_PS_PP038L.Offset = i;
                    oDS_PS_PP038L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP038L.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim());         //선택
                    oDS_PS_PP038L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("InputGbn").Value.ToString().Trim());      //투입구분
                    oDS_PS_PP038L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());      //품목코드
                    oDS_PS_PP038L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());      //품목이름
                    oDS_PS_PP038L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("ItemGpCd").Value.ToString().Trim());      //품목그룹
                    oDS_PS_PP038L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim());      //배치번호
                    oDS_PS_PP038L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("PartNo").Value.ToString().Trim());        //PartNo
                    oDS_PS_PP038L.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("Weight").Value.ToString().Trim());        //중량
                    oDS_PS_PP038L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("Unit").Value.ToString().Trim());          //단위
                    oDS_PS_PP038L.SetValue("U_ColDt01", i, oRecordSet01.Fields.Item("DueDate").Value.ToString("yyyyMMdd"));     //납기요구일
                    oDS_PS_PP038L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("CntcCode").Value.ToString().Trim());      //사번
                    oDS_PS_PP038L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("CntcName").Value.ToString().Trim());      //이름
                    oDS_PS_PP038L.SetValue("U_ColDt02", i, oRecordSet01.Fields.Item("CGDate").Value.ToString("yyyyMMdd"));      //청구일자
                    oDS_PS_PP038L.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("ProcType").Value.ToString().Trim());      //조달방식
                    oDS_PS_PP038L.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("ImportYN").Value.ToString().Trim());      //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                    oDS_PS_PP038L.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("EmergYN").Value.ToString().Trim());       //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                    oDS_PS_PP038L.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("RCode").Value.ToString().Trim());         //재청구사유
                    oDS_PS_PP038L.SetValue("U_ColReg21", i, oRecordSet01.Fields.Item("RName").Value.ToString().Trim());         //재청구사유내용
                    oDS_PS_PP038L.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("Comments").Value.ToString().Trim());      //비고
                    oDS_PS_PP038L.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());      //DocEntry
                    oDS_PS_PP038L.SetValue("U_ColNum02", i, oRecordSet01.Fields.Item("LineId").Value.ToString().Trim());        //LineId
                    oDS_PS_PP038L.SetValue("U_ColNum03", i, oRecordSet01.Fields.Item("VisOrder").Value.ToString().Trim());      //VisOrder
                    oDS_PS_PP038L.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("Object").Value.ToString().Trim());        //Object
                    oDS_PS_PP038L.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("LogInst").Value.ToString().Trim());       //LogInst
                    oDS_PS_PP038L.SetValue("U_ColNum04", i, oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());       //LineNum

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private bool PS_PP038_AddData()
        {
            bool functionReturnValue = false;
            short loopCount;
            string sQry;
            string InputGbn;       //투입구분
            string ItemCode;       //품목코드
            string ItemName;       //품목이름
            string ItemGpCd;       //품목그룹
            string BatchNum;       //배치번호
            string PartNo;         //PartNo
            string Weight;         //중량
            string Unit;           //단위
            string DueDate;        //납기요구일
            string CntcCode;       //사번
            string CntcName;       //이름
            string CGDate;         //청구일자
            string ProcType;       //조달방식
            string ImportYN;       //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
            string EmergYN;        //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
            string RCode;          //재청구사유(2018.09.17 송명규, 김석태 과장 요청)
            string RName;          //재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
            string Comments;       //비고
            string DocEntry;       //DocEntry
            string LineId;         //LineID
            string VisOrder;       //VisOrder
            string Object_Renamed; //Object
            string LogInst;        //LogInst
            string LineNum;        //U_LineNum
            string MainEntry;      //작업지시문서번호
            string BPLId;          //사업장코드
            string FullOrdNum;     //전체작번
            string errMessage = string.Empty;
            string errPoint = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                MainEntry = oForm.Items.Item("MainEntry").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
                FullOrdNum = oForm.Items.Item("FullOrdNum").Specific.Value.ToString().Trim();
                oMat01.FlushToDataSource();
                
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_PP038L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        InputGbn =oDS_PS_PP038L.GetValue("U_ColReg02", loopCount).ToString().Trim();       //투입구분
                        ItemCode =oDS_PS_PP038L.GetValue("U_ColReg03", loopCount).ToString().Trim();       //품목코드
                        ItemName =oDS_PS_PP038L.GetValue("U_ColReg04", loopCount).ToString().Trim();       //품목이름
                        ItemGpCd =oDS_PS_PP038L.GetValue("U_ColReg05", loopCount).ToString().Trim();       //품목그룹
                        BatchNum =oDS_PS_PP038L.GetValue("U_ColReg06", loopCount).ToString().Trim();       //배치번호
                        PartNo = oDS_PS_PP038L.GetValue("U_ColReg07", loopCount).ToString().Trim();        //PartNo
                        Weight = oDS_PS_PP038L.GetValue("U_ColQty01", loopCount).ToString().Trim();        //중량
                        Unit = oDS_PS_PP038L.GetValue("U_ColReg08", loopCount).ToString().Trim();          //단위
                        DueDate = oDS_PS_PP038L.GetValue("U_ColDt01", loopCount).ToString().Trim();        //납기요구일
                        CntcCode = oDS_PS_PP038L.GetValue("U_ColReg10", loopCount).ToString().Trim();      //사번
                        CntcName = oDS_PS_PP038L.GetValue("U_ColReg11", loopCount).ToString().Trim();      //이름
                        CGDate = oDS_PS_PP038L.GetValue("U_ColDt02", loopCount).ToString().Trim();         //청구일자
                        ProcType = oDS_PS_PP038L.GetValue("U_ColReg13", loopCount).ToString().Trim();      //조달방식
                        ImportYN = oDS_PS_PP038L.GetValue("U_ColReg15", loopCount).ToString().Trim();      //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                        EmergYN = oDS_PS_PP038L.GetValue("U_ColReg16", loopCount).ToString().Trim();       //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                        RCode = oDS_PS_PP038L.GetValue("U_ColReg20", loopCount).ToString().Trim();         //재청구사유(2018.09.17 송명규, 김석태 과장 요청)
                        RName = oDS_PS_PP038L.GetValue("U_ColReg21", loopCount).ToString().Trim();         //재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
                        Comments = oDS_PS_PP038L.GetValue("U_ColReg14", loopCount).ToString().Trim();      //비고
                        DocEntry = oDS_PS_PP038L.GetValue("U_ColNum01", loopCount).ToString().Trim();      //DocEntry
                        LineId = oDS_PS_PP038L.GetValue("U_ColNum02", loopCount).ToString().Trim();        //LineID
                        VisOrder = oDS_PS_PP038L.GetValue("U_ColSum03", loopCount).ToString().Trim();      //VisOrder
                        Object_Renamed = oDS_PS_PP038L.GetValue("U_ColReg18", loopCount).ToString().Trim();//Object
                        LogInst = oDS_PS_PP038L.GetValue("U_ColReg19", loopCount).ToString().Trim();       //LogInst
                        LineNum = oDS_PS_PP038L.GetValue("U_ColNum04", loopCount).ToString().Trim();       //LineNum

                        sQry = "            EXEC [PS_PP038_03] ";
                        sQry += "'" + InputGbn + "',";        //투입구분
                        sQry += "'" + ItemCode + "',";        //품목코드
                        sQry += "'" + ItemName + "',";        //품목이름
                        sQry += "'" + ItemGpCd + "',";        //품목그룹
                        sQry += "'" + BatchNum + "',";        //배치번호
                        sQry += "'" + PartNo + "',";          //PartNo
                        sQry += "'" + Weight + "',";          //중량
                        sQry += "'" + Unit + "',";            //단위
                        sQry += "'" + DueDate + "',";         //납기요구일
                        sQry += "'" + CntcCode + "',";        //사번
                        sQry += "'" + CntcName + "',";        //이름
                        sQry += "'" + CGDate + "',";          //청구일자
                        sQry += "'" + ProcType + "',";        //조달방식
                        sQry += "'" + ImportYN + "',";        //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                        sQry += "'" + EmergYN + "',";         //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                        sQry += "'" + RCode + "',";           //재청구사유(2018.09.17 송명규, 김석태 과장 요청)
                        sQry += "'" + RName + "',";           //재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
                        sQry += "'" + Comments + "',";        //비고
                        sQry += "'" + DocEntry + "',";        //DocEntry
                        sQry += "'" + LineId + "',";          //LineID
                        sQry += "'" + VisOrder + "',";        //VisOrder
                        sQry += "'" + Object_Renamed + "',";  //Object
                        sQry += "'" + LogInst + "',";         //LogInst
                        sQry += "'" + LineNum + "',";         //LineNum
                        sQry += "'" + MainEntry + "',";       //선택한 작업지시문서번호
                        sQry += "'" + BPLId + "',";           //사업장코드
                        sQry += "'" + FullOrdNum + "',";      //전체작번
                        sQry += "'" + PSH_Globals.oCompany.UserSignature + "'";   //UserSign

                        if (PS_PP038_Check_DupReq(MainEntry, ItemCode, LineId) == true)
                        {
                            if (oMat01.Columns.Item("RCode").Cells.Item(loopCount + 1).Specific.Selected == null)
                            {
                                errMessage = loopCount + 1 + "행의 원재료 청구가 중복되어 재청구사유를 필수로 입력하여야 합니다. 등록이 취소되었습니다.";
                                errPoint = "RCode";
                                throw new Exception();
                            }
                        }
                        oRecordSet01.DoQuery(sQry);
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                    if (errPoint != string.Empty)
                    {
                        oForm.Items.Item(errPoint).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    errPoint = string.Empty;
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DPS_PP038_DeleteData
        /// </summary>
        private bool PS_PP038_DeleteData()
        {
            bool functionReturnValue = false;
            short loopCount ;
            string sQry;
            string DocEntry;//문서번호
            string LineId;  //라인번호
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다";
                    throw new Exception();
                }
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_PP038L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {

                        DocEntry = oDS_PS_PP038L.GetValue("U_ColNum01", loopCount).ToString().Trim(); //문서번호
                        LineId = oDS_PS_PP038L.GetValue("U_ColNum02", loopCount).ToString().Trim();   //라인번호

                        sQry = "            EXEC [PS_PP038_05] "; //구매견적까지 진행된 구매요청이 존재하는지 체크
                        sQry += "'" + DocEntry + "',";      //문서번호
                        sQry += "'" + LineId + "'";         //라인번호

                        oRecordSet01.DoQuery(sQry);

                        sQry = "";
                        //구매견적진행되지 않은 건만 삭제 가능
                        if (oRecordSet01.Fields.Item("CNT").Value > 0)
                        {
                           errMessage =  loopCount + 1 + "행(" + DocEntry + "-" + LineId + ")은 원자재 구매청구가 등록되었습니다. 삭제할 수 없습니다. 삭제명령은 중단됩니다.";
                        }
                        sQry = "            EXEC [PS_PP038_04] ";
                        sQry += "'" + DocEntry + "',"; //문서번호
                        sQry += "'" + LineId + "'";    //라인번호
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if(errMessage != null)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP038_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP038'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP038_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            string OrdNum;
            string OrdSub1;
            string OrdSub2;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass(); 
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oUID == "Mat01")
                {
                    if (oCol == "ItemCode")
                    {
                        oDS_PS_PP038L.SetValue("U_ColReg03", oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value);
                        if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_PP038L.GetValue("U_ColReg03", oRow - 1).ToString().Trim()))
                        {
                            PS_PP038_AddMatrixRow(oRow, false);
                        }
                    }
                    else if (oCol == "CntcCode")
                    {
                        oMat01.FlushToDataSource();
                        oDS_PS_PP038L.SetValue("U_ColReg11", oRow - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value + "'",""));
                        oMat01.LoadFromDataSource();
                    }
                    oMat01.AutoResizeColumns();
                }
                else if (oUID == "CntcCode")
                {
                    oForm.Items.Item("CntcName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(oUID).Specific.Value + "'","");
                }
                else if (oUID == "OrdNum" | oUID == "OrdSub1" | oUID == "OrdSub2")
                {
                    OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
                    OrdSub1 = oForm.Items.Item("OrdSub1").Specific.Value.ToString().Trim();
                    OrdSub2 = oForm.Items.Item("OrdSub2").Specific.Value.ToString().Trim();

                    sQry = "        SELECT      CASE";
                    sQry += "                 WHEN T0.U_JakMyung = '' THEN (SELECT FrgnName FROM OITM WHERE ItemCode = T0.U_ItemCode)";
                    sQry += "                 ELSE T0.U_JakMyung";
                    sQry += "             END AS [JakMyung],";
                    sQry += "             CASE";
                    sQry += "                 WHEN T0.U_JakSize = '' THEN (SELECT U_Size FROM OITM WHERE ItemCode = T0.U_ItemCode)";
                    sQry += "                 ELSE T0.U_JakSize";
                    sQry += "             END AS [JakSize]";
                    sQry += " FROM        [@PS_PP020H] AS T0";
                    sQry += " WHERE       T0.U_JakName = '" + OrdNum + "'";
                    sQry += "             AND T0.U_SubNo1 = CASE WHEN '" + OrdSub1 + "' = '' THEN '00' ELSE '" + OrdSub1 + "' END";
                    sQry += "             AND T0.U_SubNo2 = CASE WHEN '" + OrdSub2 + "' = '' THEN '000' ELSE '" + OrdSub2 + "' END";

                    oRecordSet01.DoQuery(sQry);

                    oForm.Items.Item("JakMyung").Specific.Value = oRecordSet01.Fields.Item("JakMyung").Value;
                    oForm.Items.Item("JakSize").Specific.Value = oRecordSet01.Fields.Item("JakSize").Value;
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP038_CheckBeforeSearch() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP038_MTX01();
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
                        if (PS_PP038_AddData() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PS_PP038_MTX02(oForm.Items.Item("GridRow").Specific.Value);
                        PS_PP038_AddMatrixRow(oMat01.RowCount, false);

                    }
                    else if (pVal.ItemUID == "BtnDel")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
                        {
                            if (PS_PP038_DeleteData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP038_MTX02(oForm.Items.Item("GridRow").Specific.Value);
                            PS_PP038_AddMatrixRow(oMat01.RowCount, false);
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Mat01" & pVal.ColUID == "Check" & pVal.Row > 0)
                    {
                        //빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
                        if (oMat01.RowCount >= pVal.Row)
                        {
                            if (PS_PP038_CheckOKYN(pVal.Row) == true)
                            {
                                errMessage = "해당 자재는 구매요청승인이 이미 처리되어 선택할 수 없습니다";
                                oMat01.Columns.Item("Check").Cells.Item(pVal.Row).Specific.Checked = false;
                                throw new Exception();
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
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string OrdGbn;
            string InputGbn;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            OrdGbn = oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim();
                            InputGbn = oMat01.Columns.Item("InputGbn").Cells.Item(pVal.Row).Specific.Selected.Value;

                            PS_SM021 oTempClass = new PS_SM021(); //품목선택 화면 호출
                            oTempClass.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, OrdGbn, InputGbn, oForm.Items.Item("BPLID").Specific.Selected.Value.ToString().Trim());
                            BubbleEvent = false;
                            return;
                        }
                        else if (pVal.ColUID == "CntcCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CntcCode");  //사번 포맷서치설정
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
                            PS_PP038_MTX02(pVal.Row);
                            PS_PP038_AddMatrixRow(oMat01.RowCount, false);
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
                        if (pVal.ItemUID == "Mat01")
                        {
                            PS_PP038_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }
                        else
                        {
                            PS_PP038_FlushToItemValue(pVal.ItemUID, 0 ,"");
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
                    PS_PP038_AddMatrixRow(oMat01.VisualRowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP038L);
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
                    PS_PP038_FormResize();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
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
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
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
