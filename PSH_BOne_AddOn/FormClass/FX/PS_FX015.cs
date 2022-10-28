using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 고정자산 분개처리
    /// </summary>
    internal class PS_FX015 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_FX015H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_FX015L; //등록라인

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX015.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_FX015_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_FX015");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_FX015_CreateItems();
                PS_FX015_ComboBox_Setting();
                PS_FX015_FormClear();

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", true); // 복제
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
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
        /// 화면 Item 생성
        /// </summary>
        private void PS_FX015_CreateItems()
        {
            try
            {
                oDS_PS_FX015H = oForm.DataSources.DBDataSources.Item("@PS_FX015H");
                oDS_PS_FX015L = oForm.DataSources.DBDataSources.Item("@PS_FX015L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_FX015_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("CLTCOD").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", dataHelpClass.User_BPLID(), false, false);
                oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                sQry = " SELECT U_Minor, U_CdName";
                sQry += "  FROM [@PS_SY001L]";
                sQry += " WHERE Code = 'FX005'";

                oForm.Items.Item("FxHisCls").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("FxHisCls").Specific, sQry, "%", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }
        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_FX015_HeaderSpaceLineDel()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_FX015H.GetValue("U_CLTCOD", 0)) || string.IsNullOrEmpty(oDS_PS_FX015H.GetValue("U_StdYM", 0)))
                {
                    errMessage = "사업장, 년월은 필수입력 사항입니다.확인하세요.";
                    throw new Exception();
                }
                ReturnValue = true;
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
            return ReturnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_FX015_MatrixSpaceLineDel()
        {
            bool ReturnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                }
                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_FX015L.Offset = i;
                        if (string.IsNullOrEmpty(oDS_PS_FX015L.GetValue("U_AcctCode", i)))
                        {
                            errMessage = "계정과목코드가 없습니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_FX015L.GetValue("U_AcctName", i)))
                        {
                            errMessage = "계정과목명이 없습니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
                }
                if (string.IsNullOrEmpty(oDS_PS_FX015L.GetValue("U_AcctCode", oMat01.VisualRowCount - 1)))
                {
                    oDS_PS_FX015L.RemoveRecord(oMat01.VisualRowCount - 1);
                }
                oMat01.LoadFromDataSource();
                ReturnValue = true;
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
            return ReturnValue;
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_FX015_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;
                    oForm.Items.Item("Btn03").Enabled = true;
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("StdYM").Enabled = true;
                    oForm.Items.Item("FxHisCls").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;
                    oForm.Items.Item("Btn03").Enabled = true;
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("StdYM").Enabled = true;
                    oForm.Items.Item("FxHisCls").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("StdYM").Enabled = false;
                    oForm.Items.Item("FxHisCls").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    if (string.IsNullOrEmpty(oForm.Items.Item("JdtCC").Specific.Value))
                    {
                        oForm.Items.Item("JdtDate").Enabled = true;
                        oForm.Items.Item("Btn02").Enabled = true;
                        oForm.Items.Item("Btn03").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("JdtDate").Enabled = false;
                        if (oForm.Items.Item("JdtCC").Specific.Value.ToString().Trim() == "Y")
                        {
                            oForm.Items.Item("Btn02").Enabled = false;
                            oForm.Items.Item("Btn03").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("Btn02").Enabled = false;
                            oForm.Items.Item("Btn03").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_FX015_MTX01
        /// </summary>
        private void PS_FX015_MTX01()
        {
            string errMessage = string.Empty;
            int i; 
            string sQry;
            string CLTCOD; //사업장
            string StdYM; //기준년월
            string FxHisCls; //구분
            double totalDebit = 0; //차변계
            double totalCredit = 0; //대변계
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
                FxHisCls = oForm.Items.Item("FxHisCls").Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_FX015_01] '";
                sQry += CLTCOD + "','";
                sQry += StdYM + "','";
                sQry += FxHisCls + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_FX015L.Clear();

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    oDS_PS_FX015H.SetValue("U_DebitT", 0, "0"); //차변계
                    oDS_PS_FX015H.SetValue("U_CreditT", 0, "0"); //대변계
                    throw new Exception();
                }

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_FX015L.Size)
                    {
                        oDS_PS_FX015L.InsertRecord(i);
                    }
                    oMat01.AddRow();
                    oDS_PS_FX015L.Offset = i;
                    oDS_PS_FX015L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_FX015L.SetValue("U_CLTCOD", i, oRecordSet.Fields.Item("CLTCOD").Value.ToString().Trim()); //사업장
                    oDS_PS_FX015L.SetValue("U_ShortCD", i, oRecordSet.Fields.Item("ShortCD").Value.ToString().Trim()); //GL계정
                    oDS_PS_FX015L.SetValue("U_ShortNM", i, oRecordSet.Fields.Item("ShortNM").Value.ToString().Trim()); //GL계정명
                    oDS_PS_FX015L.SetValue("U_AcctCode", i, oRecordSet.Fields.Item("AcctCode").Value.ToString().Trim()); //관리계정
                    oDS_PS_FX015L.SetValue("U_AcctName", i, oRecordSet.Fields.Item("AcctName").Value.ToString().Trim()); //관리계정명
                    oDS_PS_FX015L.SetValue("U_Debit", i, oRecordSet.Fields.Item("Debit").Value.ToString().Trim()); //차변
                    oDS_PS_FX015L.SetValue("U_Credit", i, oRecordSet.Fields.Item("Credit").Value.ToString().Trim()); //대변
                    oDS_PS_FX015L.SetValue("U_ProfCode", i, oRecordSet.Fields.Item("ProfCode").Value.ToString().Trim()); //배부규칙
                    oDS_PS_FX015L.SetValue("U_ProfName", i, oRecordSet.Fields.Item("ProfName").Value.ToString().Trim()); //배부규칙명
                    oDS_PS_FX015L.SetValue("U_LineMemo", i, oRecordSet.Fields.Item("LineMemo").Value.ToString().Trim()); //적요6
                    oDS_PS_FX015L.SetValue("U_VatBP", i, oRecordSet.Fields.Item("VatBP").Value.ToString().Trim()); //거래처
                    oDS_PS_FX015L.SetValue("U_VatBPNM", i, oRecordSet.Fields.Item("VatBPNM").Value.ToString().Trim()); //거래처명
                    oDS_PS_FX015L.SetValue("U_VatBPRgN", i, oRecordSet.Fields.Item("VatBPRgN").Value.ToString().Trim()); //사업자등록번호

                    totalDebit += oRecordSet.Fields.Item("Debit").Value;
                    totalCredit += oRecordSet.Fields.Item("Credit").Value;

                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }

                oDS_PS_FX015H.SetValue("U_DebitT", 0, Convert.ToString(totalDebit)); //차변계
                oDS_PS_FX015H.SetValue("U_CreditT", 0, Convert.ToString(totalCredit)); //대변계

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_FX015_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FX015'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oDS_PS_FX015H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_FX015H.SetValue("DocEntry", 0, DocEntry);
                }
                oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oDS_PS_FX015H.SetValue("U_StdYM", 0, DateTime.Now.ToString("yyyyMM"));
                oDS_PS_FX015H.SetValue("U_JdtDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                
                oForm.Items.Item("FxHisCls").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 분개 DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PS_FX015_Create_oJournalEntries(short ChkType)
        {
            bool returnValue = false;
            int i;
            int j;
            int RetVal;
            int errCode = 0;
            int errDiCode = 0;
            double SDebit;
            double SCredit;
            string ShortName;
            string SAcctCode;
            string SPrcCode;
            string SLineMemo;
            string sDocDate;
            string sTransId = string.Empty;
            string sCC;
            string sQry;
            string ErrLine = string.Empty;
            string errDiMsg = string.Empty;
            SAPbobsCOM.JournalEntries f_oJournalEntries = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                j = 1;

                sDocDate = oDS_PS_FX015H.GetValue("U_JdtDate", 0).ToString(); //일자의 문자열 포맷(yyyy-MM-dd) 확인 필요

                f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                f_oJournalEntries.ReferenceDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null); //전기일
                f_oJournalEntries.DueDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);
                f_oJournalEntries.TaxDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);
                f_oJournalEntries.TransactionCode = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                f_oJournalEntries.UserFields.Fields.Item("U_BPLId").Value = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    ShortName = oMat01.Columns.Item("ShortCD").Cells.Item(i).Specific.Value;
                    SAcctCode = oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.Value; //관리계정
                    SDebit = Convert.ToDouble(oMat01.Columns.Item("Debit").Cells.Item(i).Specific.Value); //차변
                    SCredit = Convert.ToDouble(oMat01.Columns.Item("Credit").Cells.Item(i).Specific.Value); //차변
                    SPrcCode = oMat01.Columns.Item("ProfCode").Cells.Item(i).Specific.Value; //배부규칙

                    if (Convert.ToString(oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value).Length > 50)
                    {
                        errCode = 7;
                        ErrLine = oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value;
                        throw new Exception();
                    }
                    SLineMemo = oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value.ToString().Trim();
                    f_oJournalEntries.Lines.Add();

                    if (!string.IsNullOrEmpty(SAcctCode))
                    {
                        f_oJournalEntries.Lines.SetCurrentLine(j - 1);
                        f_oJournalEntries.Lines.AccountCode = SAcctCode; //관리계정
                        f_oJournalEntries.Lines.ShortName = ShortName; //G/L계정/BP 코드
                        f_oJournalEntries.Lines.LineMemo = SLineMemo; //적요
                        f_oJournalEntries.Lines.CostingCode = SPrcCode; //배부규칙
                        f_oJournalEntries.Lines.Debit = SDebit; //차변
                        f_oJournalEntries.Lines.Credit = SCredit; //대변
                        if (SAcctCode == "21109010")
                        {
                            f_oJournalEntries.Lines.TaxGroup = "A8";
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_With").Value = "Y";
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_VATYN").Value = "Y";
                        }
                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_BPLId").Value = oMat01.Columns.Item("CLTCOD").Cells.Item(i).Specific.Value.ToString().Trim();
                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatBP").Value = oMat01.Columns.Item("VatBP").Cells.Item(i).Specific.Value;
                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatBPName").Value = oMat01.Columns.Item("VatBPNM").Cells.Item(i).Specific.Value;
                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatRegN").Value = oMat01.Columns.Item("VatBPRgN").Cells.Item(i).Specific.Value;
                        j += 1;
                    }
                }

                RetVal = f_oJournalEntries.Add();//완료
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
                }

                sCC = "Y";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "Update [@PS_FX015H] Set U_JdtNo = '" + sTransId + "', U_JdtDate = '" + sDocDate + "', U_JdtCC = '" + sCC + "' ";
                    sQry = sQry + "Where DocEntry = '" + oDS_PS_FX015H.GetValue("DocEntry", 0).ToString().Trim() + "'";
                    oRecordSet.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PS_FX015H.SetValue("U_JdtNo", 0, sTransId);
                oDS_PS_FX015H.SetValue("U_JdtCC", 0, sCC);

                oForm.Items.Item("Btn02").Enabled = false;
                oForm.Items.Item("Btn03").Enabled = true;

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("적요는 50글자 초과 등록 불가합니다. (" + ErrLine + "번째 라인)", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }

            return returnValue;
        }

        /// <summary>
        /// 분개취소 DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PS_FX015_Cancel_oJournalEntries(short ChkType)
        {
            bool returnValue = false;
            int errCode = 0;
            int errDiCode = 0;
            int RetVal;
            string sCC;
            string sQry;
            string errDiMsg = string.Empty;
            string sTransId = string.Empty;
            SAPbobsCOM.JournalEntries f_oJournalEntries = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                if (f_oJournalEntries.GetByKey(Convert.ToInt32(oDS_PS_FX015H.GetValue("U_JdtNo", 0).ToString().Trim())) == false)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
                }

                RetVal = f_oJournalEntries.Cancel();//완료
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 2;
                    throw new Exception();
                }

                sCC = "N";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "  Update [@PS_FX015H] Set U_JdtCanNo = '" + sTransId + "', U_JdtCC = '" + sCC + "' ";
                    sQry += " Where DocEntry = '" + oDS_PS_FX015H.GetValue("DocEntry", 0).ToString().Trim() + "'";
                    oRecordSet.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PS_FX015H.SetValue("U_JdtCanNo", 0, sTransId);
                oDS_PS_FX015H.SetValue("U_JdtCC", 0, sCC);

                oForm.Items.Item("Btn02").Enabled = false;
                oForm.Items.Item("Btn03").Enabled = false;

                PSH_Globals.SBO_Application.StatusBar.SetText("분개취소 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("취소할 분개번호 조회 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }

            return returnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_FX015_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false; // BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
                                return;
                            }
                            if (PS_FX015_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        PS_FX015_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value))
                            {
                                errMessage = "분개처리일을 먼저 입력하세요.";
                                throw new Exception();
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                errMessage = "문서가 Close 또는 Cancel 되었습니다.";
                                throw new Exception();
                            }
                            else
                            {
                                if (PS_FX015_Create_oJournalEntries(1) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            errMessage = "먼저 저장한 후 분개 처리 바랍니다.";
                            throw new Exception();
                        }
                        PS_FX015_FormItemEnabled();
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value))
                            {
                                errMessage = "분개처리일을 먼저 입력하세요.";
                                throw new Exception();
                            }
                            else if (oForm.Items.Item("JdtCC").Specific.Value != "Y")
                            {
                                errMessage = "분개생성:Y일 때 취소 할 수 있습니다.";
                                throw new Exception();
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                errMessage = "문서가 Close 또는 Cancel 되었습니다.";
                                throw new Exception();
                            }
                            else
                            {
                                if (PS_FX015_Cancel_oJournalEntries(1) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            errMessage = "먼저 저장한 후 분개 처리 바랍니다.";
                            throw new Exception();
                        }
                        PS_FX015_FormItemEnabled();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE  && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == false)
                        {
                            PS_FX015_FormItemEnabled();
                        }
                    }
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
                BubbleEvent = false;
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PS_FX015_FormItemEnabled();
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX015H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX015L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                            oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("DocEntry").Specific.Value = "";
                            oForm.Items.Item("StdYM").Specific.Value = "";
                            PS_FX015_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_FX015_FormItemEnabled();
                            PS_FX015_FormClear();
                            oForm.Items.Item("StdYM").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_FX015_FormItemEnabled();
                            break;
                        case "1287": //복제
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
