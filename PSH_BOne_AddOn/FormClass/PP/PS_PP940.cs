using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 11.일일생산계획등록
    /// </summary>
    internal class PS_PP940 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP940H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP940L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oDocEntry01;

        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP940.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP940_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP940");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP940_CreateItems();
                PS_PP940_ComboBox_Setting();
                PS_PP940_EnableMenus();
                PS_PP940_SetDocument(oFromDocEntry01);
                PS_PP940_FormResize();
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
        private void PS_PP940_CreateItems()
        {
            try
            {
                oDS_PS_PP940H = oForm.DataSources.DBDataSources.Item("@PS_PP940H");
                oDS_PS_PP940L = oForm.DataSources.DBDataSources.Item("@PS_PP940L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP940_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                //// 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP940_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.EnableMenu("1283", true);
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_PP940_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PS_PP940_FormItemEnabled();
                    PS_PP940_AddMatrixRow(0, true);
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP940_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP940_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    ////각모드에따른 아이템설정
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    PS_PP940_FormClear();           ////UDO방식
                    oForm.EnableMenu("1281", true); ////찾기
                    oForm.EnableMenu("1282", false); ////추가

                    oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    ////각모드에따른 아이템설정
                    oForm.Items.Item("Code").Specific.VALUE = "";
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.EnableMenu("1281", false); ////찾기
                    oForm.EnableMenu("1282", true);  ////추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    ////각모드에따른 아이템설정
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
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
        /// PS_PP940_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP940_AddMatrixRow(int oRow, bool RowIserted = false)
        {
            try
            {
                oForm.Freeze(true);
                ////행추가여부
                if (RowIserted == false)
                {
                    oDS_PS_PP940L.InsertRecord((oRow));
                }
                oMat01.AddRow();
                oDS_PS_PP940L.Offset = oRow;
                oDS_PS_PP940L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// PS_PP940_MTX01
        /// </summary>
        private void PS_PP940_MTX01()
        {
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string errMessage = null;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                Param01 = oForm.Items.Item("Param01").Specific.VALUE.ToString().Trim();
                Param02 = oForm.Items.Item("Param01").Specific.VALUE.ToString().Trim();
                Param03 = oForm.Items.Item("Param01").Specific.VALUE.ToString().Trim();
                Param04 = oForm.Items.Item("Param01").Specific.VALUE.ToString().Trim();

                oForm.Freeze(true);
                Query01 = "SELECT 10";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                }
                ProgressBar01.Text = "조회시작";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP940L.InsertRecord((i));
                    }
                    oDS_PS_PP940L.Offset = i;
                    oDS_PS_PP940L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP940L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
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
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP940_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP940'", "");
                if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
                {
                    oForm.Items.Item("Code").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("Code").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        ///// <summary>
        ///// PS_PP084_DI_API01
        ///// </summary>
        ///// <returns></returns>
        //private bool PS_PP940_DI_API()
        //{
        //    bool returnValue = true;
        //    int i;
        //    int j = 0;
        //    int RetVal;
        //    string errCode = string.Empty;
        //    int ResultDocNum = 0;
        //    string errDiMsg = string.Empty;
        //    int errDiCode = 0;
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        //    SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
        //    try
        //    {
        //        if (PSH_Globals.oCompany.InTransaction == true)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        PSH_Globals.oCompany.StartTransaction();
        //        oMat01.FlushToDataSource();

        //        oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
        //        oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.VALUE));
        //        oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.VALUE);
        //        oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("InDate").Specific.VALUE, "&&&&-&&-&&"));

        //        for (i = 0; i <= ItemInformationCount - 1; i++)
        //        {
        //            if (ItemInformation[i].Check == true)
        //            {
        //                goto Continue_First;
        //            }
        //            if (i != 0)
        //            {
        //                oDIObject.Lines.Add();
        //            }
        //            oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //            oDIObject.Lines.WarehouseCode = Strings.Trim(oForm.Items.Item("WhsCode").Specific.VALUE);
        //            oDIObject.Lines.BaseType = Convert.ToInt32("22");
        //            oDIObject.Lines.BaseEntry = ItemInformation[i].OPORNo;
        //            oDIObject.Lines.BaseLine = ItemInformation[i].POR1No;
        //            for (j = i; j <= Information.UBound(ItemInformation); j++)
        //            {
        //                if (ItemInformation[j].Check == true)
        //                {
        //                    goto Continue_Second;
        //                }
        //                if ((ItemInformation[i].ItemCode != ItemInformation[j].ItemCode | ItemInformation[i].OPORNo != ItemInformation[j].OPORNo | ItemInformation[i].POR1No != ItemInformation[j].POR1No))
        //                {
        //                    goto Continue_Second;
        //                }
        //                ////같은것
        //                oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation[j].Quantity;
        //                oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[j].BatchNum;
        //                oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[j].Quantity;
        //                oDIObject.Lines.BatchNumbers.Add();
        //                ItemInformation[j].PDN1No = LineNumCount;
        //                ItemInformation[j].Check = true;
        //            Continue_Second:

        //        }
        //            LineNumCount = LineNumCount + 1;
        //        Continue_First:

        //    }
        //        RetVal = oDIObject.Add();

        //        if (RetVal != 0)
        //        {
        //            PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
        //            errCode = "1";
        //            throw new Exception();
        //        }
        //        if (PSH_Globals.oCompany.InTransaction == true)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //            ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
        //            oForm.Items.Item("OIGNNo").Specific.VALUE = ResultDocNum;
        //            oDS_PS_PP048H.SetValue("U_OIGNNo", 0, Convert.ToString(ResultDocNum));
        //        }
        //        oMat01.LoadFromDataSource();
        //        oMat01.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        returnValue = false;
        //        if (PSH_Globals.oCompany.InTransaction)
        //        {
        //            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        if (errCode == "1")
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
        //    }

        //    return returnValue;
        //}

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP940_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errMessage = null;
            try
            {
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    PS_PP940_FormClear();
                }
                //사업장 미입력 시
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Selected.VALUE))
                {
                    errMessage = "사업장이 선택되지 않았습니다.";
                    throw new Exception();
                }

                //등록일자 미입력 시
                if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.VALUE))
                {
                    errMessage = "등록일자를 선택하지 않았습니다.";
                    throw new Exception();
                }
                //라인정보 미입력 시
                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                oMat01.FlushToDataSource();
                oDS_PS_PP940L.RemoveRecord(oDS_PS_PP940L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP940_FormClear();
                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
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
        /// PS_PP940_MatrixColumnSetting
        /// </summary>
        /// <returns></returns>
        private void PS_PP940_MatrixColumnSetting(string pYM)
        {
            string Dt = null;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            try
            {
                Dt = DateTime.Parse(pYM).AddMonths(1).AddDays(-1).ToString("yyyyMMdd");
                //해당 월의 마지막 날

                int LastDay = 0;
                LastDay = Convert.ToInt16(codeHelpClass.Right(Dt, 2));

                int loopCount = 0;

                for (loopCount = 1; loopCount <= oMat01.Columns.Count - 1; loopCount++)
                {
                    oMat01.Columns.Item(loopCount).Editable = true;
                }

                int DisableColumn = 0;
                string DisableColumnString = null;

                string Temp = null;
                string DayName = null;
                DayOfWeek DayNum;
                

                for (loopCount = 1; loopCount <= LastDay; loopCount++)
                {
                    DayNum = DateTime.Parse(codeHelpClass.Left(pYM, 0) + loopCount.ToString().PadLeft(2, '0')).DayOfWeek;

                    //DayNum = Convert.ToString(DateAndTime.WeekDay(Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Left(pYM, 6) + Temp + Convert.ToString(loopCount), "&&&&-&&-&&"))));
                    switch (DayNum)
                    {
                        case DayOfWeek.Sunday:
                            DayName = "일";
                            break;
                        case DayOfWeek.Monday:
                            DayName = "월";
                            break;
                        case DayOfWeek.Tuesday:
                            DayName = "화";
                            break;
                        case DayOfWeek.Wednesday:
                            DayName = "수";
                            break;
                        case DayOfWeek.Thursday:
                            DayName = "목";
                            break;
                        case DayOfWeek.Friday:
                            DayName = "금";
                            break;
                        case DayOfWeek.Saturday:
                            DayName = "토";
                            break;
                    }
                    DisableColumnString = "D" + Temp + Convert.ToString(loopCount);
                    oMat01.Columns.Item(DisableColumnString).TitleObject.Caption = Convert.ToString(loopCount) + "일(" + DayName + ")";

                    if (DayName == "일")
                    {
                        oMat01.Columns.Item(DisableColumnString).BackColor = Information.RGB(255, 0, 0);
                        //빨간색
                    }
                    else if (DayName == "토")
                    {
                        oMat01.Columns.Item(DisableColumnString).BackColor = Information.RGB(0, 128, 255);
                        //하늘색

                    }
                    else
                    {
                        oMat01.Columns.Item(DisableColumnString).BackColor = Information.RGB(255, 255, 255);
                        //흰색
                    }

                }


                if (LastDay != 31)
                {
                    DisableColumn = 31 - LastDay;
                    for (loopCount = 0; loopCount <= DisableColumn - 1; loopCount++)
                    {
                        DisableColumnString = "D" + Convert.ToString(31 - loopCount);
                        oMat01.Columns.Item(DisableColumnString).Editable = false;
                        //해당월의 말일이 존재하지 않으면 막음
                    }

                }
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {

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

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "PS_PP940")
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP940_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            ////해야할일 작업
                            oDocEntry01 = oForm.Items.Item("Code").Specific.VALUE;
                            oFormMode01 = oForm.Mode;

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP940_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            ////해야할일 작업
                            oDocEntry01 = oForm.Items.Item("Code").Specific.VALUE;
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ItemUID == "PS_PP940")
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
                            if (pVal.ItemUID == "1")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    if (pVal.ActionSuccess == true)
                                    {
                                        PS_PP940_FormItemEnabled();
                                        PS_PP940_AddMatrixRow(0, true);
                                        ////UDO방식일때
                                    }
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    if (pVal.ActionSuccess == true)
                                    {
                                        PS_PP940_FormItemEnabled();
                                    }
                                }
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
            object ChildForm01 = null;
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "ItmMsort")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItmMsort").Cells.Item(pVal.Row).Specific.VALUE))
                            {
                                dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ItmMsort");
                                //매트릭스의 품목중분류 포맷서치 적용
                            }
                        }
                        else if (pVal.ColUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.VALUE))
                            {
                                //PS_SM010 PS_SM010 = new PS_SM010();
                                //PS_SM010.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                //BubbleEvent = false;
                            }
                        }
                        else if (pVal.ColUID == "MachCode")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("MachCode").Cells.Item(pVal.Row).Specific.VALUE))
                            {
                                dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode");
                                //매트릭스의 설비코드 포맷서치 적용
                            }
                        }
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.VALUE))
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmBsort", "");
                        //품목대분류 포맷서치 적용
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdNum")
                        {
                            PS_PP030 PS_PP030 = new PS_PP030();
                            PS_PP030.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                            BubbleEvent = false;

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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {


                        if ((pVal.ItemUID == "Mat01"))
                        {

                            if (pVal.ColUID == "ItmMsort")
                            {
                                oDS_PS_PP940L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                                oDS_PS_PP940L.SetValue("U_ItmMname", pVal.Row - 1, dataHelpClass.Get_ReData("U_CodeName", "U_Code", "[@PSH_ITMMSORT]", "'" + oDS_PS_PP940L.GetValue("U_ItmMsort", pVal.Row - 1).ToString().Trim() + "'", ""));

                                if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP940L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP940_AddMatrixRow(pVal.Row);
                                }
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {

                                oDS_PS_PP940L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                                oDS_PS_PP940L.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oDS_PS_PP940L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim() + "'",""));
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (pVal.ColUID == "MachCode")
                            {
                                oDS_PS_PP940L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.VALUE);
                                oDS_PS_PP940L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P009' And U_Minor = '" + oDS_PS_PP940L.GetValue("U_MachCode", pVal.Row - 1).ToString().Trim() + "'", 0, 1));

                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        else
                        {
                            if ((pVal.ItemUID == "ItmBsort"))
                            {
                                oDS_PS_PP940H.SetValue("U_ItmBName", 0, dataHelpClass.Get_ReData("Name", "Code", "[@PSH_ITMBSORT]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'",""));
                            }
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {

                    if (pVal.ItemUID == "YM")
                    {
                        if (pVal.ItemChanged == true)
                        {

                            //매트릭스 컬럼 설정
                            //PS_PP940_MatrixColumnSetting(oForm.Items.Item("YM").Specific.VALUE);

                        }
                    }

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
                    PS_PP940_FormResize();
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
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        {
            int i;
            try
            {
                if ((oLastColRow01 > 0))
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_PP940L.RemoveRecord(oDS_PS_PP940L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP940_AddMatrixRow(0);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_PP940L.GetValue("U_ItmMsort", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP940_AddMatrixRow(oMat01.RowCount);
                            }
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP940_FormItemEnabled();
                            ////UDO방식
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP940_FormItemEnabled();
                            ////UDO방식
                            PS_PP940_AddMatrixRow(0, true);
                            ////UDO방식
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                                     //레코드이동버튼
                            PS_PP940_FormItemEnabled();
                            //PS_PP940_MatrixColumnSetting(oForm.Items.Item("YM").Specific.VALUE);
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
    }
}
