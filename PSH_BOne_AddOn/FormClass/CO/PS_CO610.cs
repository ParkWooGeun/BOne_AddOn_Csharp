using System;
using System.Collections.Generic;

using SAPbouiCOM;

using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 고정자산 본계정 대체
	/// </summary>
	internal class PS_CO610 : PSH_BaseClass
	{
		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		//public SAPbouiCOM.Form oForm02;
		public SAPbouiCOM.Matrix oMat01;
			
		private SAPbouiCOM.DBDataSource oDS_PS_CO610H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO610L; //등록라인
		
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO610.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO610_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO610");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);
				oForm.EnableMenu("1293", true);

                CreateItems();
                ComboBox_Setting();
                CF_ChooseFromList();
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
				PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                oDS_PS_CO610H = oForm.DataSources.DBDataSources.Item("@PS_CO610H");
                oDS_PS_CO610L = oForm.DataSources.DBDataSources.Item("@PS_CO610L");

                oMat01 = oForm.Items.Item("Mat01").Specific; //매트릭스 데이터 셋
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oDS_PS_CO610H.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
                oDS_PS_CO610H.SetValue("U_JdtDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                //사업장 리스트
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
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
        /// Combobox 설정
        /// </summary>
        private void ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                sQry = "  SELECT    Code,";
                sQry += "           Name";
                sQry += " FROM      [@PSH_Account_list]";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("AcctCode"), sQry, "", "");

                sQry = "  SELECT    distinct";
                sQry += "           b.U_Minor,";
                sQry += "           b.U_CdName";
                sQry += " FROM      [@PS_SY001H] AS a";
                sQry += "           Inner Join";
                sQry += "           [@PS_SY001L] AS b";
                sQry += "               On a.Code = b.Code";
                sQry += "               And a.Code = 'FX007'";
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("LongYear"), sQry, "", "");
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
        /// ChooseFromList
        /// </summary>
        private void CF_ChooseFromList()
        {
            try
            {

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Initial Setting
        /// </summary>
        private void Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();            

            try
            {    
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //사업장
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
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Amt").Enabled = false;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oMat01.Columns.Item("Check").Editable = true;
                    oMat01.Columns.Item("PrcCode").Editable = true;
                    oMat01.Columns.Item("LineMemo").Editable = true;
                    oMat01.Columns.Item("AcctCode").Editable = true;
                    oMat01.Columns.Item("Number1").Editable = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oMat01.Columns.Item("LongYear").Editable = true;
                    oForm.Items.Item("YM").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    if (oForm.Items.Item("JdtCC").Specific.VALUE == "Y")
                    {
                        oForm.Items.Item("Amt").Enabled = false;
                        oForm.Items.Item("JdtDate").Enabled = false;
                        oMat01.Columns.Item("Check").Editable = false;
                        oMat01.Columns.Item("PrcCode").Editable = false;
                        oMat01.Columns.Item("LineMemo").Editable = false;
                        oMat01.Columns.Item("AcctCode").Editable = false;
                        oMat01.Columns.Item("Number1").Editable = false;
                        oForm.Items.Item("YM").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("JdtDate").Enabled = false;
                        oForm.Items.Item("Comments").Enabled = false;
                        oForm.Items.Item("Btn02").Enabled = false;
                        oMat01.Columns.Item("LongYear").Editable = false;
                    }
                    else if (oForm.Items.Item("JdtCC").Specific.VALUE == "N")
                    {
                        oForm.Items.Item("Amt").Enabled = false;
                        oForm.Items.Item("JdtDate").Enabled = false;
                        oMat01.Columns.Item("Check").Editable = false;
                        oMat01.Columns.Item("PrcCode").Editable = false;
                        oMat01.Columns.Item("LineMemo").Editable = false;
                        oMat01.Columns.Item("AcctCode").Editable = false;
                        oMat01.Columns.Item("Number1").Editable = false;
                        oForm.Items.Item("YM").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("JdtDate").Enabled = false;
                        oForm.Items.Item("Comments").Enabled = false;
                        oForm.Items.Item("Btn02").Enabled = false;
                        oMat01.Columns.Item("LongYear").Editable = false;
                        oForm.Items.Item("Btn03").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Amt").Enabled = true;
                        oForm.Items.Item("JdtDate").Enabled = true;
                        oMat01.Columns.Item("Check").Editable = true;
                        oMat01.Columns.Item("PrcCode").Editable = true;
                        oMat01.Columns.Item("LineMemo").Editable = true;
                        oMat01.Columns.Item("AcctCode").Editable = true;
                        oMat01.Columns.Item("LongYear").Editable = true;
                        oMat01.Columns.Item("Number1").Editable = true;
                        oForm.Items.Item("YM").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("JdtDate").Enabled = true;
                        oForm.Items.Item("Btn02").Enabled = true;
                        oForm.Items.Item("Btn03").Enabled = true;
                    }

                    oForm.Items.Item("DocNum").Enabled = false;
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
                        oDS_PS_CO610L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oMat01.LoadFromDataSource();
                        break;
                    case 1:
                        oDS_PS_CO610L.InsertRecord(oRow);
                        oDS_PS_CO610L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oMat01.LoadFromDataSource();
                        break;
                }
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
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO610'", "");

                if (Convert.ToDouble(DocNum) == 0)
                {
                    oDS_PS_CO610H.SetValue("DocNum", 0, "1");
                }
                else
                {
                    oDS_PS_CO610H.SetValue("DocNum", 0, DocNum);
                }
            }
            catch (Exception ex)
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
            string YM;
            string BPLId;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);

                YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                sQry = "EXEC [PS_CO610_01] '" + BPLId + "','" + YM + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_CO610L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO610L.Size)
                    {
                        oDS_PS_CO610L.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PS_CO610L.Offset = i;
                    oDS_PS_CO610L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO610L.SetValue("U_DocDate", i, Convert.ToDateTime(oRecordSet01.Fields.Item("DocDate").Value).ToString("yyyyMMdd"));
                    oDS_PS_CO610L.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_CardName", i, oRecordSet01.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_Price", i, oRecordSet01.Fields.Item("Price").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_Desc", i, oRecordSet01.Fields.Item("Desc").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_LineMemo", i, oRecordSet01.Fields.Item("Desc").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_Qty", i, oRecordSet01.Fields.Item("Qty").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_Unit", i, oRecordSet01.Fields.Item("Unit").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_CntcName", i, oRecordSet01.Fields.Item("CntcName").Value.ToString().Trim());
                    oDS_PS_CO610L.SetValue("U_UseDept", i, oRecordSet01.Fields.Item("UseDept").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.Columns.Item("Desc").Visible = false;
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
                if(oDS_PS_CO610H.GetValue("U_BPLId", 0) == "" || oDS_PS_CO610H.GetValue("U_YM", 0) == "")
                {
                    errCode = "1";
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장, 년월은 필수입력 사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (oMat01.VisualRowCount < 1)
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_CO610L.Offset = i;
                    }
                }

                if (string.IsNullOrEmpty(oDS_PS_CO610L.GetValue("U_CardCode", oMat01.VisualRowCount - 1)))
                {
                    oDS_PS_CO610L.RemoveRecord(oMat01.VisualRowCount - 1);
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
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            
            return returnValue;
        }

        /// <summary>
        /// 오류 메시지 일괄 처리
        /// </summary>
        /// <param name="ErrNum">오류번호</param>
        private void Item_Error_Message(short ErrNum)
        {
            try
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("분개처리일을 먼저 입력하세요.");
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.MessageBox("문서가 Close 또는 Cancel 되었습니다.");
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("분개생성:Y일 때 취소 할 수 있습니다.");
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.MessageBox("거래처코드와, 사업장을 먼저 입력하세요.");
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.MessageBox("대체계정 필드에 값이 입력되지 않았습니다.");
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.MessageBox("배부규칙 필드에 값이 입력되지 않았습니다..");
                }
            }
            catch(Exception ex)
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
            int j;
            int errCode = 0;
            string errDiMsg = string.Empty;
            int errDiCode = 0;
            string ErrLine = string.Empty;
            int RetVal;
            double SDebit;
            double SCredit;
            string SAcctCode;
            string SPrcCode;
            string SLineMemo;
            string SVatBP;
            string sDocDate;
            string sTransId = string.Empty;
            string sCC;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset); ;
            SAPbobsCOM.JournalEntries f_oJournalEntries = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                j = 1;

                sDocDate = oDS_PS_CO610H.GetValue("U_JdtDate", 0).ToString(); //일자의 문자열 포맷(yyyy-MM-dd) 확인 필요

                f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                f_oJournalEntries.ReferenceDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null); //전기일
                f_oJournalEntries.DueDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);
                f_oJournalEntries.TaxDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == "True")
                    {
                        SAcctCode = oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.Value; //관리계정
                        SDebit = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value; //BP코드
                        SPrcCode = oMat01.Columns.Item("PrcCode").Cells.Item(i).Specific.Value; //배부규칙
                        SVatBP = oMat01.Columns.Item("CardCode").Cells.Item(i).Specific.Value.ToString().Trim(); //거래처코드

                        if (Convert.ToString(oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value).Length > 50)
                        {
                            errCode = 7;
                            ErrLine = oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value;
                            throw new Exception();
                        }

                        SLineMemo = oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value.ToString().Trim(); //적요

                        f_oJournalEntries.Lines.Add();

                        if (!string.IsNullOrEmpty(SAcctCode))
                        {
                            f_oJournalEntries.Lines.SetCurrentLine(j - 1);
                            f_oJournalEntries.Lines.AccountCode = SAcctCode; //관리계정
                            f_oJournalEntries.Lines.ShortName = SAcctCode; //G/L계정/BP 코드
                            f_oJournalEntries.Lines.LineMemo = SLineMemo; //적요
                            f_oJournalEntries.Lines.CostingCode = SPrcCode; //배부규칙
                            f_oJournalEntries.Lines.Debit = SDebit; //차변
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatBP").Value = SVatBP; //거래처코드
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatBPName").Value = oMat01.Columns.Item("CardName").Cells.Item(i).Specific.Value.ToString().Trim(); //거래처명
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_VatRegN").Value = dataHelpClass.Get_ReData("VATRegNum", "CardCode", "[OCRD]", "'" + SVatBP + "'", ""); //사업자등록번호
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_OcrName").Value = dataHelpClass.Get_ReData("PrcName", "PrcCode", "[OPRC]", "'" + SPrcCode + "'", ""); //배부규칙명
                            f_oJournalEntries.Lines.UserFields.Fields.Item("U_Number1").Value = oMat01.Columns.Item("Number1").Cells.Item(i).Specific.Value.ToString().Trim(); //관리항목
                            f_oJournalEntries.UserFields.Fields.Item("U_BPLId").Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                            j = j + 1;
                        }
                    }
                }

                SCredit = Convert.ToDouble(oForm.Items.Item("Amt").Specific.Value);

                f_oJournalEntries.Lines.Add();
                f_oJournalEntries.Lines.SetCurrentLine(j - 1);
                f_oJournalEntries.Lines.AccountCode = "59109020";//관리계정
                f_oJournalEntries.Lines.ShortName = "59109020";//G/L계정/BP 코드
                //f_oJournalEntries.Lines.ContraAct = "55102010"
                f_oJournalEntries.Lines.LineMemo = "본계정에 대체"; //적요
                f_oJournalEntries.Lines.Credit = SCredit; //차변
                f_oJournalEntries.UserFields.Fields.Item("U_BPLId").Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장

                //완료
                RetVal = f_oJournalEntries.Add();
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
                    sQry = "  Update [@PS_CO610H] Set U_JdtNo = '" + sTransId + "', U_JdtDate = '" + sDocDate + "', U_JdtCC = '" + sCC + "' ";
                    sQry += " Where DocNum = '" + oDS_PS_CO610H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PS_CO610H.SetValue("U_JdtNo", 0, sTransId);
                oDS_PS_CO610H.SetValue("U_JdtCC", 0, sCC);

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 분개취소 DI
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool Cancel_oJournalEntries(short ChkType)
        {
            bool returnValue = false;
            
            //int i = 0;
            //short ErrNum = 0;
            int errCode = 0;
            int errDiCode = 0;
            string errDiMsg = null;
            int RetVal = 0;
            string sTransId = null;

            //string SCardCode = null;
            //string sDocDate = null;
            string sCC = null;
            string sQry = null;

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

                if (f_oJournalEntries.GetByKey(Convert.ToInt32(oDS_PS_CO610H.GetValue("U_JdtNo", 0).ToString().Trim())) == false)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = 1;
                    throw new Exception();
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
                    sQry = "  Update [@PS_CO610H] Set U_JdtCanNo = '" + sTransId + "', U_JdtCC = '" + sCC + "' ";
                    sQry += " Where DocNum = '" + oDS_PS_CO610H.GetValue("DocNum", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                
                oDS_PS_CO610H.SetValue("U_JdtCanNo", 0, sTransId);
                oDS_PS_CO610H.SetValue("U_JdtCC", 0, sCC);

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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }
            
            return returnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        private void PS_CO610_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocNum").Specific.Value.ToString().Trim();

                WinTitle = "[PS_CO610_01] 고정자산 본계정 대체";
                ReportName = "PS_CO610_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry)); //문서번호
                
                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }







        #region Raise_ItemEvent
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			short i = 0;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			object TempForm01 = null;
        //			short ErrNum = 0;
        //			int SumTot = 0;
        //			string CheckYN = null;

        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			//// 객체 정의 및 데이터 할당

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.EventType) {

        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						if (pval.ItemUID == "1") {
        //							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //								//UPGRADE_WARNING: oMat01.Columns(CardCode).Cells(1).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (!string.IsNullOrEmpty(oMat01.Columns.Item("CardCode").Cells.Item(1).Specific.VALUE)) {
        //									if (HeaderSpaceLineDel() == false) {
        //										BubbleEvent = false;
        //										// BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
        //										return;
        //									}
        //									if (MatrixSpaceLineDel() == false) {
        //										BubbleEvent = false;
        //										return;
        //									}
        //								} else {
        //									SubMain.Sbo_Application.MessageBox("자료 조회 후 추가하세요.");
        //									BubbleEvent = false;
        //									return;
        //								}
        //							}
        //						} else if (pval.ItemUID == "Prt") {
        //							PS_CO610_Print_Report01();

        //						//// 상각자료 불러오기
        //						} else if (pval.ItemUID == "Btn01") {
        //							MTX01();
        //						//// DI API - 분개 생성
        //						} else if (pval.ItemUID == "Btn02") {

        //							//대체 계정 및 배부규칙 체크 S
        //							for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //								//UPGRADE_WARNING: oMat01.Columns(Check).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == "True") {
        //									//UPGRADE_WARNING: oMat01.Columns(AcctCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if (string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.VALUE)) {
        //										ErrNum = 5;
        //										Item_Error_Message(ref ErrNum);
        //										BubbleEvent = false;
        //										return;
        //									}

        //									//UPGRADE_WARNING: oMat01.Columns(PrcCode).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if (string.IsNullOrEmpty(oMat01.Columns.Item("PrcCode").Cells.Item(i).Specific.VALUE)) {
        //										ErrNum = 6;
        //										Item_Error_Message(ref ErrNum);
        //										BubbleEvent = false;
        //										return;
        //									}
        //								}
        //							}


        //							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //								//UPGRADE_WARNING: oForm01.Items(JdtDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm01.Items.Item("JdtDate").Specific.VALUE)) {
        //									ErrNum = 1;
        //									Item_Error_Message(ref ErrNum);
        //									BubbleEvent = false;
        //									return;
        //									//UPGRADE_WARNING: oForm01.Items(Status).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								} else if (oForm01.Items.Item("Status").Specific.VALUE == "C") {
        //									ErrNum = 2;
        //									Item_Error_Message(ref ErrNum);
        //									BubbleEvent = false;
        //									return;
        //								} else {
        //									if (Create_oJournalEntries(ref 1) == false) {
        //										BubbleEvent = false;
        //										return;
        //									}
        //								}
        //								FormItemEnabled();
        //							} else {
        //								SubMain.Sbo_Application.MessageBox("저장후 분개확정하세요.");
        //								BubbleEvent = false;
        //								return;
        //							}
        //							//
        //						//// DI API - 분개 취소
        //						} else if (pval.ItemUID == "Btn03") {
        //							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //								//UPGRADE_WARNING: oForm01.Items(JdtDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if (string.IsNullOrEmpty(oForm01.Items.Item("JdtDate").Specific.VALUE)) {
        //									ErrNum = 1;
        //									Item_Error_Message(ref ErrNum);
        //									BubbleEvent = false;
        //									return;
        //									//UPGRADE_WARNING: oForm01.Items(JdtCC).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								} else if (oForm01.Items.Item("JdtCC").Specific.VALUE != "Y") {
        //									ErrNum = 3;
        //									Item_Error_Message(ref ErrNum);
        //									BubbleEvent = false;
        //									return;
        //									//UPGRADE_WARNING: oForm01.Items(Status).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								} else if (oForm01.Items.Item("Status").Specific.VALUE == "C") {
        //									ErrNum = 2;
        //									Item_Error_Message(ref ErrNum);
        //									BubbleEvent = false;
        //									return;
        //								} else {
        //									if (Cancel_oJournalEntries(ref 1) == false) {
        //										BubbleEvent = false;
        //										return;
        //									}
        //								}
        //							} else {
        //								MDC_Com.MDC_GF_Message(ref "먼저 저장한 후 분개 처리 바랍니다.", ref "W");
        //								BubbleEvent = false;
        //								return;
        //							}
        //							//
        //						} else if (pval.ItemUID == "Mat01") {
        //							if (pval.ColUID == "Check") {
        //								//선택된 중량 체크 S
        //								for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //									//UPGRADE_WARNING: oMat01.Columns(Check).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									if (oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked == "True") {
        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										SumTot = SumTot + oMat01.Columns.Item("Price").Cells.Item(i).Specific.VALUE;
        //									}
        //								}

        //								//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oForm01.Items.Item("Amt").Specific.VALUE = SumTot;
        //								//선택된 중량 체크 E

        //							}

        //						} else {
        //							//                If pval.ItemChanged = True Then
        //							//
        //							//                End If
        //						}
        //						break;


        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						break;
        //					// 거래처코드


        //					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //						////5
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CLICK:
        //						////6
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //						////7
        //						//            If pval.ItemChanged = True Then
        //						//            End If
        //						//
        //						if (pval.ColUID == "Check") {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							CheckYN = oMat01.Columns.Item("Check").Cells.Item(1).Specific.Checked;
        //							for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //								if (Convert.ToBoolean(CheckYN) == false) {
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked = "True";
        //								} else {
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("Check").Cells.Item(i).Specific.Checked = "False";
        //								}
        //							}
        //						}
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
        //						oLast_Item_UID = pval.ItemUID;
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						break;
        //				}

        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						break;
        //					//
        //					//             ' 저장 후 추가 가능처리
        //					//                If pval.ItemUID = "1" Then
        //					//                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
        //					//                          oForm01.Mode = fm_OK_MODE
        //					//                          Call Sbo_Application.ActivateMenuItem("1282")
        //					//                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
        //					//                        FormItemEnabled
        //					//                        AddMatrixRow 1, oMat01.RowCount, True
        //					//                    End If
        //					//                End If
        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						if (pval.Action_Success == true) {
        //							oSeq = 1;
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
        //						break;

        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //						////11
        //						break;
        //					//                AddMatrixRow 1, oMat01.VisualRowCount, True
        //					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //						////18
        //						if (oSeq == 1) {
        //							oSeq = 0;
        //						}
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
        //						oLast_Item_UID = pval.ItemUID;
        //						break;

        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						SubMain.RemoveForms(oFormUniqueID01);
        //						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oForm01 = null;
        //						//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oMat01 = null;
        //						break;
        //				}
        //			}

        //			return;
        //			Raise_ItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //						//행닫기
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
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1281":
        //						//찾기
        //						FormItemEnabled();
        //						break;
        //					//                oForm01.Items("ItemCode").Click ct_Regular
        //					case "1282":
        //						//추가
        //						FormItemEnabled();
        //						FormClear();
        //						AddMatrixRow(0, oMat01.RowCount, ref true);
        //						//풀어야함.
        //						oForm01.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //						break;

        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						FormItemEnabled();
        //						if (oMat01.VisualRowCount > 0) {
        //							//UPGRADE_WARNING: oMat01.Columns(AcctCode).Cells(oMat01.VisualRowCount).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (!string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(oMat01.VisualRowCount).Specific.VALUE)) {
        //								//AddMatrixRow 1, oMat01.RowCount, True 풀어야함.
        //							}
        //						}
        //						break;
        //					case "1293":
        //						//행닫기
        //						break;

        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        //						////34 - 추가
        //						break;
        //					//            FormClear
        //					//            If Create_oJournalEntries(2) = False Then
        //					//                BubbleEvent = False
        //					//                Exit Sub
        //					//            End If
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35 - 업데이트
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
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if ((eventInfo.BeforeAction == true)) {
        //				////작업
        //			} else if ((eventInfo.BeforeAction == false)) {
        //				////작업
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion









    }
}
