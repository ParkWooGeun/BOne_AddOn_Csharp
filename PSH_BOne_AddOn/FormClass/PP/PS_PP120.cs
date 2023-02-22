using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// PSMT임가공내역등록 및 전기
    /// </summary>
    internal class PS_PP120 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP120H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP120L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP120.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP120_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP120");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP120_CreateItems();
                PS_PP120_ComboBox_Setting();
                PS_PP120_EnableMenus();
                PS_PP120_SetDocument(oFormDocEntry);
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
        private void PS_PP120_CreateItems()
        {
            try
            {
                oDS_PS_PP120H = oForm.DataSources.DBDataSources.Item("@PS_PP120H");
                oDS_PS_PP120L = oForm.DataSources.DBDataSources.Item("@PS_PP120L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");
                oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");

                oForm.DataSources.UserDataSources.Item("FrDate").ValueEx = DateTime.Now.ToString("yyyyMM01");
                oForm.DataSources.UserDataSources.Item("ToDate").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                oForm.Items.Item("Acct").Specific.Value = "83102010";

                oForm.Items.Item("empty").Visible = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP120_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId=2 order by BPLId", "2", false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP120_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제") //행삭제전 행삭제가능여부검사
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) //추가,수정모드일때행삭제가능검사
                    {
                        if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value)))//새로추가된 행인경우, 삭제하여도 무방하다
                        {
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oForm.Items.Item("ApDocNum").Specific.Value))
                            {
                                errMessage = "AP전기된문서입니다. 행삭제를 할수 없습니다.";
                                throw new Exception();
                            }
                            else if (oForm.Items.Item("Canceled").Specific.Value == "Y")
                            {
                                errMessage = "취소된문서는 수정할수 없습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (!string.IsNullOrEmpty(oForm.Items.Item("ApDocNum").Specific.Value))
                    {
                        errMessage = "AP전기된문서입니다. 취소할수 없습니다.";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("Canceled").Specific.Value == "Y")
                    {
                        errMessage = "이미취소된문서입니다.";
                        throw new Exception();
                    }
                }
                returnValue = true;
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
            return returnValue;
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP120_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true,true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP120_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP120_FormItemEnabled();
                    PS_PP120_AddMatrixRow(0, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP120_FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP120_FormClear();
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("empty").Click();
                    oForm.Items.Item("Mat01").Enabled = true; //활성 메트릭스
                    oForm.Items.Item("CardCode").Enabled = true; //활성 고객코드
                    oForm.Items.Item("InDate").Enabled = true; //활성 작성일
                    oForm.Items.Item("BPLId").Enabled = true; //활성 사업장
                    oForm.Items.Item("Btn1").Enabled = false; //비활성 전기버튼
                    oForm.Items.Item("DocEntry").Enabled = false; //비활성 문서번호
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Btn1").Enabled = false; //전기버튼 비활성화
                    oForm.Items.Item("DocEntry").Enabled = true; //문서번호활성화
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("InDate").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    if (!string.IsNullOrEmpty(oDS_PS_PP120H.GetValue("U_ApDocNum", 0).Replace(" ", "")))
                    {
                        oForm.Items.Item("Btn1").Enabled = false;
                        oForm.Items.Item("ApDate").Enabled = false;
                        oForm.Items.Item("Acct").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("InDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                    }
                    else if (oDS_PS_PP120H.GetValue("Canceled", 0) == "Y")
                    {
                        oForm.Items.Item("Btn1").Enabled = false;
                        oForm.Items.Item("Acct").Enabled = false;
                        oForm.Items.Item("ApDate").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("InDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Btn1").Enabled = true;
                        oForm.Items.Item("Acct").Enabled = true;
                        oForm.Items.Item("ApDate").Enabled = true;
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("CardCode").Enabled = true;
                        oForm.Items.Item("InDate").Enabled = true;
                        oForm.Items.Item("BPLId").Enabled = true;
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
        /// 
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP120_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)//행추가여부
                {
                    oDS_PS_PP120L.InsertRecord((oRow));
                }
                oMat01.AddRow();
                oDS_PS_PP120L.Offset = oRow;
                oDS_PS_PP120L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP120L.SetValue("U_ReHour", oRow, "0");
                oMat01.LoadFromDataSource();
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
        /// PS_PP120_MTX01
        /// </summary>
        private void PS_PP120_MTX01()
        {
            string errMessage = string.Empty; 
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();

                Query01 = "SELECT 10";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                ProgressBar01.Text = "조회중";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP120L.InsertRecord((i));
                    }
                    oDS_PS_PP120L.Offset = i;
                    oDS_PS_PP120L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP120L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP120_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP120'", "");
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PP120_DocTotalSum
        /// </summary>
        /// <returns></returns>
        private bool PS_PP120_DocTotalSum()
        {
            bool returnValue = false;
            int i;
            double sDocTotal = 0;

            try
            {
                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    sDocTotal += Convert.ToDouble(oMat01.Columns.Item("Total").Cells.Item(i).Specific.Value);
                }
                oForm.Items.Item("DocTotal").Specific.Value = sDocTotal;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
            return returnValue;
        }

        /// <summary>
        /// PS_PP120_DocTotalSum
        /// </summary>
        /// <returns></returns>
        private void PS_PP120_UpdateToPP120H()
        {
            string Query01;
            string DocEntry;
            string sApDocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oDS_PS_PP120H.GetValue("DocEntry", 0).ToString().Trim();
                sApDocNum = oDS_PS_PP120H.GetValue("U_ApDocNum", 0).ToString().Trim();

                Query01 = "UPDATE";
                Query01 += " [@PS_PP120H]";
                Query01 += " SET";
                Query01 += " U_ApDocNum = '" + sApDocNum + "'";
                Query01 += " , U_ApDate = '" + dataHelpClass.ConvertDateType(oDS_PS_PP120H.GetValue("U_ApDate", 0),"") + "'";
                Query01 += " WHERE DocEntry = '" + DocEntry + "'";

                oRecordSet01.DoQuery(Query01);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                oForm.Items.Item("empty").Click();
                oForm.Items.Item("Btn1").Enabled = false;
                oForm.Items.Item("Acct").Enabled = false;
                oForm.Items.Item("ApDate").Enabled = false;
                oForm.Items.Item("CardCode").Enabled = false;
                oForm.Items.Item("InDate").Enabled = false;
                oForm.Items.Item("BPLId").Enabled = false;
                oForm.Items.Item("Mat01").Enabled = false;

                PSH_Globals.SBO_Application.SetStatusBarMessage("처리 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP120_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("ApDate").Specific.Value))
                {
                    errMessage = "전기일자는 필수입니다.";
                    ClickCode = "ApDate";
                    type = "F";
                    throw new Exception();
                }
                else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("ApDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일자를 확인하고, 회계부서로 문의하세요.";
                    type = "F";
                    ClickCode = "ApDate";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                {
                    errMessage = "거래처코드는 필수입니다.";
                    ClickCode = "CardCode";
                    type = "F";
                    throw new Exception();
                }
                else if (oForm.Items.Item("DocTotal").Specific.Value == 0)
                {
                    errMessage = "총계가 0 입니다.";
                    ClickCode = "DocTotal";
                    type = "F";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    type = "M";
                    throw new Exception();
                }
                else
                {
                    //값이 한줄들어있을때 한줄삭제후 갱신한다거나한다면
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("PorNum").Cells.Item(1).Specific.Value))
                    {
                        errMessage = "Matrix값이 한줄이상은 있어야합니다.";
                        ClickCode = "PorNum";
                        type = "M";
                        throw new Exception();
                    }

                }
                for (i = 1; i <= (oMat01.VisualRowCount - 1); i++)
                {
                    if ((string.IsNullOrEmpty(oMat01.Columns.Item("PorNum").Cells.Item(i).Specific.Value)))
                    {
                        errMessage = "작지번호는 필수입니다.";
                        ClickCode = "PorNum";
                        type = "M";
                        throw new Exception();
                    }
                    else if ((string.IsNullOrEmpty(oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value)))
                    {
                        errMessage = "공정코드는 필수입니다.";
                        ClickCode = "CpCode";
                        type = "M";
                        throw new Exception();
                    }
                    else if ((string.IsNullOrEmpty(oMat01.Columns.Item("AskDate").Cells.Item(i).Specific.Value)))
                    {
                        errMessage = "지원일자는 필수입니다.";
                        ClickCode = "AskDate";
                        type = "M";
                        throw new Exception();
                    }
                }

                oDS_PS_PP120L.RemoveRecord(oDS_PS_PP120L.Size - 1);
                oMat01.LoadFromDataSource();
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    PS_PP120_FormClear();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (type == "M")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
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
            return returnValue;
        }


        private bool PS_PP120_DI_API()
        {
            bool returnValue = false;
            int i;
            int RetVal;
            int ResultDocNum;
            int errDiCode = 0;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
                oDIObject.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service; //서비스송장
                oDIObject.BPL_IDAssignedToInvoice = 2; //동래기본박아넣음 'oForm.Items("BPLId").Specific.Selected.Value
                oDIObject.CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim(); //공급업체
                oDIObject.CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim(); //공급처이름
                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("ApDate").Specific.Value, "yyyyMMdd", null); //전기일
                oDIObject.DocDueDate = DateTime.ParseExact(oForm.Items.Item("ApDate").Specific.Value, "yyyyMMdd", null); //만기일
                oDIObject.Comments = "기준PSMT내역문서 : " + oForm.Items.Item("DocEntry").Specific.Value;

                for (i = 0; i <= oMat01.VisualRowCount; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemDescription = oMat01.Columns.Item("PorNum").Cells.Item(i).Specific.Value; //내역에 작지번호저장
                    oDIObject.Lines.AccountCode = "83102010"; //G/L 계정 값 (외주임가공비) 83102010-외주가공비(제)
                    oDIObject.Lines.TaxCode = "V2"; //VAT그룹코드
                    oDIObject.Lines.LineTotal = oMat01.Columns.Item("Total").Cells.Item(i).Specific.Value; //공정단가로계산된총금액
                    oDIObject.Lines.UserFields.Fields.Item("U_sItemCode").Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value; //지원받은품목
                    oDIObject.Lines.UserFields.Fields.Item("U_sQty").Value = oMat01.Columns.Item("ReQty").Cells.Item(i).Specific.Value;
                }

                RetVal = oDIObject.Add();
                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    oForm.Items.Item("ApDocNum").Specific.Value = ResultDocNum;
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                returnValue = false;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
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
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP120_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP120_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if ((pVal.ItemUID == "Btn1"))
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("ApDate").Specific.Value))
                        {
                            errMessage = "전기일자는 필수입니다.";
                            ClickCode = "ApDate";
                            type = "F";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oForm.Items.Item("Acct").Specific.Value))
                        {
                            errMessage = "전기계정은 필수입니다.";
                            ClickCode = "Acct";
                            type = "F";
                            throw new Exception();
                        }

                        if (PS_PP120_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        PS_PP120_DI_API();//DI로 AP서비스송장을 생성한다.
                        PS_PP120_UpdateToPP120H();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP120_FormItemEnabled();
                                PS_PP120_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP120_FormItemEnabled();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (type == "M")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ReHour")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.MessageBox("공정코드를 먼저 선택하세요");
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }

                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //품목코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Acct", ""); //전기계정조회
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PorNum"); //작지번호
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CpCode"); //공정코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "WorkMan"); //작업자정보
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ItemCode"); //품목코드조회
                }
                else if (pVal.Before_Action == false)
                {
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false); //메트릭스 한줄선택시 반전시켜주는 구문
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (pVal.ItemUID == "1")
                        {
                            oForm.EnableMenu("1281", true); //찾기하고 다시 찾기아이콘활성화처리
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int sReHour = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PorNum")
                            {
                                //기타작업
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP120L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP120_AddMatrixRow(pVal.Row, false);
                                }
                                oDS_PS_PP120L.SetValue("U_ItemCode", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_ItemCode,U_OrdNum  FROM [@PS_PP030H] WHERE Convert(VarChar(20),U_OrdNum) + Convert(varchar(10),U_OrdSub1) + Convert(varchar(10),U_OrdSub2)='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Replace("-", "") + "'",0,1));
                                oDS_PS_PP120L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_ItemName,U_OrdNum  FROM [@PS_PP030H] WHERE Convert(VarChar(20),U_OrdNum) + Convert(varchar(10),U_OrdSub1) + Convert(varchar(10),U_OrdSub2)='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Replace("-", "") + "'", 0, 1));
                                oDS_PS_PP120L.SetValue("U_SjNo", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SjNum,U_OrdNum  FROM [@PS_PP030H] WHERE Convert(VarChar(20),U_OrdNum) + Convert(varchar(10),U_OrdSub1) + Convert(varchar(10),U_OrdSub2)='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Replace("-", "") + "'", 0, 1));
                                oDS_PS_PP120L.SetValue("U_SjTotal", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("SELECT U_SjPrice,U_OrdNum  FROM [@PS_PP030H] WHERE Convert(VarChar(20),U_OrdNum) + Convert(varchar(10),U_OrdSub1) + Convert(varchar(10),U_OrdSub2)='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Replace("-", "") + "'", 0, 1)));
                            }
                            else if (pVal.ColUID == "CpCode")
                            {
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP120L.SetValue("U_CpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName From [@PS_PP001L] WHERE U_CpCode='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'",0,1));
                                oDS_PS_PP120L.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("SELECT U_Price From [@PS_PP001L] WHERE U_CpCode='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)));
                                //작업시간을 변경할때와 같이 갱신이 이루어져야한다.
                                sReHour = oMat01.Columns.Item("ReHour").Cells.Item(pVal.Row).Specific.Value;
                                oDS_PS_PP120L.SetValue("U_Total", pVal.Row - 1, Convert.ToString(Convert.ToString(sReHour * Convert.ToDouble(oDS_PS_PP120L.GetValue("U_CpPrice", pVal.Row - 1)))));
                            }
                            else if (pVal.ColUID == "CpPrice")
                            {
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                sReHour = oMat01.Columns.Item("ReHour").Cells.Item(pVal.Row).Specific.Value;
                                oDS_PS_PP120L.SetValue("U_Total", pVal.Row - 1, Convert.ToString(Convert.ToString(sReHour * Convert.ToDouble(oDS_PS_PP120L.GetValue("U_CpPrice", pVal.Row - 1)))));
                            }
                            else if ((pVal.ColUID == "ItemCode" && !string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)))
                            {
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP120L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.GetValue("SELECT FrgnName FROM [OITM] WHERE ItemCode='" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));

                            }
                            else if (pVal.ColUID == "ReHour")
                            {
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP120L.SetValue("U_Total", pVal.Row - 1, Convert.ToString(Convert.ToString(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value * oMat01.Columns.Item("CpPrice").Cells.Item(pVal.Row).Specific.Value)));
                            }
                            else
                            {
                                oDS_PS_PP120L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if ((pVal.ItemUID == "DocEntry"))
                            {
                                oDS_PS_PP120H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if ((pVal.ItemUID == "CardCode"))
                            {
                                oDS_PS_PP120H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                            else
                            {
                                oDS_PS_PP120H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        
                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) //찾기모드에서는 계산안타기
                        {
                            PS_PP120_DocTotalSum();
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Update();
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
                    PS_PP120_FormItemEnabled();
                    PS_PP120_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP120H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP120L);
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
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;

            try
            {
                if ((oLastColRow01 > 0))
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_PP120_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i; //행을 다시 순서대로정렬해서 행번에넣고(VisualCount값은 줄어든상태)
                        }
                        oMat01.FlushToDataSource();//Matrix의 RowCount값 갯수도 줄어든 수만큼 갱신처리 해주며

                        oDS_PS_PP120L.RemoveRecord(oDS_PS_PP120L.Size - 1); //줄어든 행수만큼 DataSources값 갱신한 뒤
                        oMat01.LoadFromDataSource(); //그후 다시 데이터소스를 읽어와 화면완성을 한다.

                        //행이 없으면 한줄추가
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP120_AddMatrixRow(0, false);
                        }
                        else
                        {
                            //현재행삭제한 행의PorNum값이 있는행지우면 넘어가고 없는 마지막행값지우면 한행추가
                            if (!string.IsNullOrEmpty(oDS_PS_PP120L.GetValue("U_PorNum", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP120_AddMatrixRow(oMat01.RowCount, false);
                            }
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
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if ((PS_PP120_Validate("취소") == false))
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 취소할수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP120_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_PP120_FormItemEnabled();
                            PS_PP120_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_PP120_FormItemEnabled();
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
