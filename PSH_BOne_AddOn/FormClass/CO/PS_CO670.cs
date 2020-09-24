using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Windows.Forms;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기간비용분개등록
	/// </summary>
	internal class PS_CO670 : PSH_BaseClass
	{
		private string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		//public SAPbouiCOM.Form oForm02;
		private SAPbouiCOM.Matrix oMat01;
			
		private SAPbouiCOM.DBDataSource oDS_PS_CO670H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO670L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private int oSeq;

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO670.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO670_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO670");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

                CreateItems();
                Initial_Setting();
                FormItemEnabled();
                FormClear();
                AddMatrixRow(0, oMat01.RowCount, true);

                oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1284", true); // 취소
				oForm.EnableMenu("1293", true); // 행삭제
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
        private void CreateItems()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oDS_PS_CO670H = oForm.DataSources.DBDataSources.Item("@PS_CO670H");
                oDS_PS_CO670L = oForm.DataSources.DBDataSources.Item("@PS_CO670L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oDS_PS_CO670H.SetValue("U_StdDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PS_CO670H.SetValue("U_JdtDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                //사업장 리스트
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLID").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Initial_Setting
        /// </summary>
        private void Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("JdtDate").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("JdtDate").Enabled = true;
                }

                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormClear
        /// </summary>
        private void FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO670'", "");

                if (DocEntry == "0")
                {
                    oDS_PS_CO670H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_CO670H.SetValue("DocEntry", 0, Convert.ToString(DocEntry));
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 행추가
        /// </summary>
        /// <param name="oSeq"></param>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void AddMatrixRow(short oSeq, int oRow, bool RowIserted)
        {
            try
            {
                switch (oSeq)
                {
                    case 0:
                        oMat01.AddRow();
                        oDS_PS_CO670L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oMat01.LoadFromDataSource();
                        break;
                    case 1:
                        oDS_PS_CO670L.InsertRecord(oRow);
                        oDS_PS_CO670L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oMat01.LoadFromDataSource();
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스에 데이터 로드
        /// </summary>
        private void MTX01()
        {
            int i;
            string sQry;
            string BPLID;
            string StdDate;
            string CoAcctCD;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);

                BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
                StdDate = oForm.Items.Item("StdDate").Specific.Value.ToString().Trim();
                CoAcctCD = oForm.Items.Item("CoAcctCD").Specific.Value.ToString().Trim();

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                sQry = "EXEC [PS_CO670_01] '";
                sQry += BPLID + "','";
                sQry += StdDate + "','";
                sQry += CoAcctCD + "'";
                oRecordSet01.DoQuery(sQry);
                
                oMat01.Clear();
                oDS_PS_CO670L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }
                
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO670L.Size)
                    {
                        oDS_PS_CO670L.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PS_CO670L.Offset = i;
                    oDS_PS_CO670L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO670L.SetValue("U_BasEntry", i, oRecordSet01.Fields.Item("BasEntry").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_BasLine", i, oRecordSet01.Fields.Item("BasLine").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_AcctCode", i, oRecordSet01.Fields.Item("AcctCode").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_AcctName", i, oRecordSet01.Fields.Item("AcctName").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_Debit", i, oRecordSet01.Fields.Item("Debit").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_Credit", i, oRecordSet01.Fields.Item("Credit").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_OcrCode", i, oRecordSet01.Fields.Item("OcrCode").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_OcrName", i, oRecordSet01.Fields.Item("OcrName").Value.ToString().Trim());
                    oDS_PS_CO670L.SetValue("U_LineMemo", i, oRecordSet01.Fields.Item("LineMemo").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Update();
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Header 필수 입력 필드 체크
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errCode = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_CO670H.GetValue("U_BPLID", 0).ToString().Trim()) || string.IsNullOrEmpty(oDS_PS_CO670H.GetValue("U_StdDate", 0).ToString().Trim()))
                {
                    errCode = "1";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장, 일자는 필수입력 사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Line 필수 입력 필드 체크
        /// </summary>
        /// <returns></returns>
        private bool MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errCode = string.Empty;
            
            try
            {
                oMat01.FlushToDataSource();

                //라인
                if (oMat01.VisualRowCount <= 1)
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {

                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_CO670L.Offset = i;

                        if (string.IsNullOrEmpty(oDS_PS_CO670L.GetValue("U_AcctCode", i)))
                        {
                            errCode = "2";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_CO670L.GetValue("U_AcctName", i)))
                        {
                            errCode = "3";
                            throw new Exception();
                        }
                    }
                }

                if (string.IsNullOrEmpty(oDS_PS_CO670L.GetValue("U_AcctCode", oMat01.VisualRowCount - 1)))
                {
                    oDS_PS_CO670L.RemoveRecord(oMat01.VisualRowCount - 1);
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("계정과목코드가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("계정과목명이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            
            return returnValue;
        }

        /// <summary>
        /// 오류 메시지 일괄 처리
        /// </summary>
        /// <param name="ErrNum">오류번호</param>
        private void Item_Error_Message(short ErrNum)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (ErrNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("분개처리일을 먼저 입력하세요.", "E");
                }
                else if (ErrNum == 2)
                {
                    dataHelpClass.MDC_GF_Message("문서가 Close 또는 Cancel 되었습니다.", "E");
                }
                else if (ErrNum == 3)
                {
                    dataHelpClass.MDC_GF_Message("분개생성:Y일 때 취소 할 수 있습니다.", "E");
                }
                else if (ErrNum == 4)
                {
                    dataHelpClass.MDC_GF_Message("거래처코드와, 사업장을 먼저 입력하세요.", "E");
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 분개 DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool Create_oJournalEntries(short ChkType)
        {
            bool returnValue = false;

            int i;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            int errDiCode = 0;
            int RetVal;
            string sTransId = string.Empty;
            double SDebit;
            double SCredit;

            string SAcctCode;
            string sDocDate;
            string SPrcCode;
            string SLineMemo;

            string sCC;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.JournalEntries f_oJournalEntries = null;

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                sDocDate = oDS_PS_CO670H.GetValue("U_JdtDate", 0).ToString();

                f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                f_oJournalEntries.ReferenceDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null); //전기일
                f_oJournalEntries.DueDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);
                f_oJournalEntries.TaxDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    SAcctCode = oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.Value;
                    SDebit = Convert.ToDouble(oMat01.Columns.Item("Debit").Cells.Item(i).Specific.Value);
                    SCredit = Convert.ToDouble(oMat01.Columns.Item("Credit").Cells.Item(i).Specific.Value);
                    SPrcCode = oMat01.Columns.Item("OcrCode").Cells.Item(i).Specific.Value;
                    SLineMemo = oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value;

                    f_oJournalEntries.Lines.Add();

                    if (!string.IsNullOrEmpty(SAcctCode))
                    {
                        f_oJournalEntries.Lines.SetCurrentLine(i - 1);
                        f_oJournalEntries.Lines.AccountCode = SAcctCode; //관리계정
                        f_oJournalEntries.Lines.ShortName = SAcctCode; //G/L계정/BP 코드
                        f_oJournalEntries.Lines.LineMemo = SLineMemo; //비고

                        f_oJournalEntries.Lines.CostingCode = SPrcCode; //배부규칙
                        f_oJournalEntries.Lines.Debit = SDebit; //차변
                        f_oJournalEntries.Lines.Credit = SCredit; //대변

                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_BillCode").Value = "P90030"; //법정지출증빙코드
                        f_oJournalEntries.Lines.UserFields.Fields.Item("U_BillName").Value = "기타"; //법정지출증빙명

                        f_oJournalEntries.UserFields.Fields.Item("U_BPLId").Value = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim(); //사업장
                    }

                    //분개여부 UPDATE
                    sQry = "  UPDATE  [@PS_CO660L]";
                    sQry += " SET     U_JEYN = 'Y'";
                    sQry += " WHERE   DocEntry = " + oMat01.Columns.Item("BasEntry").Cells.Item(i).Specific.Value;
                    sQry += "         AND LineID = " + oMat01.Columns.Item("BasLine").Cells.Item(i).Specific.Value;

                    oRecordSet01.DoQuery(sQry);
                }

                //완료
                RetVal = f_oJournalEntries.Add();
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }

                sCC = "Y";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "Update [@PS_CO670H] Set U_JdtNo = '" + sTransId + "', U_JdtDate = '" + sDocDate + "', U_JdtCC = '" + sCC + "' ";
                    sQry = sQry + "Where DocNum = '" + oDS_PS_CO670H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PS_CO670H.SetValue("U_JdtNo", 0, sTransId);
                oDS_PS_CO670H.SetValue("U_JdtCC", 0, sCC);

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

                if (errCode == "1")
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 분개 DI 취소
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool Cancel_oJournalEntries(short ChkType)
        {
            bool returnValue = false;
            
            int i;
            int errCode = 0;
            int errDiCode = 0;
            string errDiMsg = string.Empty;
            int RetVal;
            string sTransId = string.Empty;
            string sCC;
            string sQry;

            SAPbobsCOM.JournalEntries f_oJournalEntries = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                if (f_oJournalEntries.GetByKey(Convert.ToInt32(oDS_PS_CO670H.GetValue("U_JdtNo", 0).ToString().Trim())) == false)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
                }

                //분개여부 환원
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    //분개여부 UPDATE
                    sQry = "        UPDATE  [@PS_CO660L]";
                    sQry = sQry + " SET     U_JEYN = 'N'";
                    sQry = sQry + " WHERE   DocEntry = " + oMat01.Columns.Item("BasEntry").Cells.Item(i).Specific.VALUE;
                    sQry = sQry + "         AND LineID = " + oMat01.Columns.Item("BasLine").Cells.Item(i).Specific.VALUE;

                    oRecordSet01.DoQuery(sQry);
                }

                //완료
                RetVal = f_oJournalEntries.Cancel();
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
                    sQry = "Update [@PS_CO670H] Set U_JdtCanNo = '" + sTransId + "', U_JdtCC = '" + sCC + "' ";
                    sQry = sQry + "Where DocNum = '" + oDS_PS_CO670H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PS_CO670H.SetValue("U_JdtCanNo", 0, sTransId);
                oDS_PS_CO670H.SetValue("U_JdtCC", 0, sCC);

                oForm.Items.Item("Btn02").Enabled = false;
                oForm.Items.Item("Btn03").Enabled = false;

                PSH_Globals.SBO_Application.StatusBar.SetText("분개취소 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                returnValue = true;
            }
            catch(Exception ex)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }
            
            return returnValue;
        }

        










        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;
        //	string sQry = null;
        //	object TempForm01 = null;
        //	short ErrNum = 0;

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	//// 객체 정의 및 데이터 할당

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true)) {
        //		switch (pval.EventType) {

        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1
        //				if (pval.ItemUID == "1") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //						if (HeaderSpaceLineDel() == false) {
        //							BubbleEvent = false;
        //							// BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
        //							return;
        //						}
        //						if (MatrixSpaceLineDel() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //				//// 상각자료 불러오기
        //				} else if (pval.ItemUID == "Btn01") {
        //					MTX01();
        //				//// DI API - 분개 생성
        //				} else if (pval.ItemUID == "Btn02") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //						//UPGRADE_WARNING: oForm01.Items(JdtDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm01.Items.Item("JdtDate").Specific.VALUE)) {
        //							ErrNum = 1;
        //							Item_Error_Message(ref ErrNum);
        //							BubbleEvent = false;
        //							return;
        //							//UPGRADE_WARNING: oForm01.Items(Status).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (oForm01.Items.Item("Status").Specific.VALUE == "C") {
        //							ErrNum = 2;
        //							Item_Error_Message(ref ErrNum);
        //							BubbleEvent = false;
        //							return;
        //						} else {
        //							if (Create_oJournalEntries(ref 1) == false) {
        //								BubbleEvent = false;
        //								return;
        //							}
        //						}

        //					} else {
        //						MDC_Com.MDC_GF_Message(ref "먼저 저장한 후 분개 처리 바랍니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					}

        //				//// DI API - 분개 취소
        //				} else if (pval.ItemUID == "Btn03") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //						//UPGRADE_WARNING: oForm01.Items(JdtDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oForm01.Items.Item("JdtDate").Specific.VALUE)) {
        //							ErrNum = 1;
        //							Item_Error_Message(ref ErrNum);
        //							BubbleEvent = false;
        //							return;
        //							//UPGRADE_WARNING: oForm01.Items(JdtCC).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (oForm01.Items.Item("JdtCC").Specific.VALUE != "Y") {
        //							ErrNum = 3;
        //							Item_Error_Message(ref ErrNum);
        //							BubbleEvent = false;
        //							return;
        //							//UPGRADE_WARNING: oForm01.Items(Status).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						} else if (oForm01.Items.Item("Status").Specific.VALUE == "C") {
        //							ErrNum = 2;
        //							Item_Error_Message(ref ErrNum);
        //							BubbleEvent = false;
        //							return;
        //						} else {
        //							if (Cancel_oJournalEntries(ref 1) == false) {
        //								BubbleEvent = false;
        //								return;
        //							}
        //						}
        //					} else {
        //						MDC_Com.MDC_GF_Message(ref "먼저 저장한 후 분개 처리 바랍니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					}

        //				} else {
        //					if (pval.ItemChanged == true) {

        //					}
        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2

        //				if (pval.ItemUID == "Mat01") {

        //					//                    If pval.ColUID = "PGCd" Then
        //					//
        //					//                        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "PGCd") '//사용자값활성
        //					//                        'Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "PGCd") '프로그램코드 포맷서치설정
        //					//
        //					//                    End If

        //				//상대계정
        //				} else if (pval.ItemUID == "CoAcctCD") {

        //					MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "CoAcctCD", "");
        //					//계정 포맷서치 설정

        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6

        //				if (pval.ItemUID == "Mat01") {
        //					if (pval.Row > 0) {
        //						oLast_Item_UID = pval.ItemUID;
        //						oLast_Col_UID = pval.ColUID;
        //						oLast_Col_Row = pval.Row;

        //						oMat01.SelectRow(pval.Row, true, false);
        //					}
        //				} else {
        //					oLast_Item_UID = pval.ItemUID;
        //					oLast_Col_UID = "";
        //					oLast_Col_Row = 0;
        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				oLast_Item_UID = pval.ItemUID;
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				break;
        //		}

        //	////BeforeAction = False
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //				////1

        //				// 저장 후 추가 가능처리
        //				if (pval.ItemUID == "1") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pval.Action_Success == true) {
        //						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //						SubMain.Sbo_Application.ActivateMenuItem("1282");
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pval.Action_Success == false) {
        //						FormItemEnabled();
        //						AddMatrixRow(1, oMat01.RowCount, ref true);
        //					}
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //				////2
        //				if (pval.Action_Success == true) {
        //					oSeq = 1;
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //				////5
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_CLICK:
        //				////6
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //				////7
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //				////8
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //				////10

        //				//상대계정
        //				if (pval.ItemUID == "CoAcctCD") {

        //					//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sQry = "SELECT AcctName FROM OACT WHERE AcctCode = '" + oForm01.Items.Item(pval.ItemUID).Specific.VALUE + "'";
        //					oRecordSet01.DoQuery(sQry);
        //					//UPGRADE_WARNING: oForm01.Items(CoAcctNM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm01.Items.Item("CoAcctNM").Specific.VALUE = Strings.Trim(oRecordSet01.Fields.Item(0).Value);

        //				}
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //				////11
        //				break;
        //			//                AddMatrixRow 1, oMat01.VisualRowCount, True
        //			case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //				////18
        //				if (oSeq == 1) {
        //					oSeq = 0;
        //				}
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //				////19
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //				////20
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //				////27
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //				////3
        //				oLast_Item_UID = pval.ItemUID;
        //				break;

        //			case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //				////4
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //				////17
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm01 = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //				break;
        //		}
        //	}

        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행닫기
        //				break;

        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;

        //		}
        //	////BeforeAction = False
        //	} else if ((pval.BeforeAction == false)) {
        //		switch (pval.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1281":
        //				//찾기
        //				FormItemEnabled();
        //				break;
        //			//                oForm01.Items("ItemCode").Click ct_Regular
        //			case "1282":
        //				//추가
        //				FormItemEnabled();
        //				FormClear();
        //				AddMatrixRow(0, oMat01.RowCount, ref true);
        //				oForm01.Items.Item("StdDate").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //				break;

        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				FormItemEnabled();
        //				if (oMat01.VisualRowCount > 0) {
        //					//UPGRADE_WARNING: oMat01.Columns(AcctCode).Cells(oMat01.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (!string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(oMat01.VisualRowCount).Specific.VALUE)) {
        //						AddMatrixRow(1, oMat01.RowCount, ref true);
        //					}
        //				}
        //				break;
        //			case "1293":
        //				//행닫기
        //				break;

        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34 - 추가
        //				break;
        //			//            FormClear
        //			//            If Create_oJournalEntries(2) = False Then
        //			//                BubbleEvent = False
        //			//                Exit Sub
        //			//            End If
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35 - 업데이트
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;

        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if ((eventInfo.BeforeAction == true)) {
        //		////작업
        //	} else if ((eventInfo.BeforeAction == false)) {
        //		////작업
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}
