using System;
using SAPbouiCOM;
using System.Text;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using System.Drawing.Imaging;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// M/G거래명세표등록
    /// </summary>
    internal class PS_PP095 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP095H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP095L; //등록라인

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP095.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP095_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP095");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP095_CreateItems();
                PS_PP095_ComboBox_Setting();
                PS_PP095_EnableMenus();
                PS_PP095_SetDocument(oFromDocEntry01);
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
        private void PS_PP095_CreateItems()
        {
            try
            {
                oDS_PS_PP095H = oForm.DataSources.DBDataSources.Item("@PS_PP095H");
                oDS_PS_PP095L = oForm.DataSources.DBDataSources.Item("@PS_PP095L");

                oMat01 = oForm.Items.Item("Mat01").Specific;

                oForm.DataSources.UserDataSources.Add("S_Weight", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("S_Weight").Specific.DataBind.SetBound(true, "", "S_Weight");

                oForm.DataSources.UserDataSources.Add("SS_Weight", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SS_Weight").Specific.DataBind.SetBound(true, "", "SS_Weight");

                oForm.DataSources.UserDataSources.Add("BoxCnt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("BoxCnt").Specific.DataBind.SetBound(true, "", "BoxCnt");

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP095_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("DocDate").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_PP095", "Gubun", "", "1", "내수");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_PP095", "Gubun", "", "2", "수출");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Gubun").Specific, "PS_PS_PP095", "Gubun" ,false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_PP095", "Way", "", "1", "편도");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_PP095", "Way", "", "2", "왕복");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Way").Specific, "PS_PS_PP095", "Way", false);

                oForm.Items.Item("TCHECK").Specific.ValidValues.Add("1", "자동확인");
                oForm.Items.Item("TCHECK").Specific.ValidValues.Add("2", "수동확인");
                oForm.Items.Item("TCHECK").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP095_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_PP095_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_PP095_FormItemEnabled();
                    PS_PP095_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP095_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PP095_Check_QRCode_PrintYN
        /// </summary>
        private string PS_PP095_Check_QRCode_PrintYN()
        {
            string functionReturnValue = "E";
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PS_PP095_10 '" + oForm.Items.Item("CardCode").Specific.VALUE + "'";

                oRecordSet01.DoQuery(sQry);

                functionReturnValue = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP095_Weight_Check
        /// </summary>
        private bool PS_PP095_Weight_Check()
        {
            bool functionReturnValue = false;
            int i;
            int result;
            int result2;
            int maxrange;
            int minrange;
            int T_Weight;
            int S_Weight = 0;
            string sQry;
            string TCHECK;
            string PackNo;
            string PackNo2;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                T_Weight = Convert.ToInt32(oDS_PS_PP095H.GetValue("U_Tweight", 0));
                TCHECK = oDS_PS_PP095H.GetValue("U_TCHECK", 0).ToString().Trim();
                PackNo2 = "0";

                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    PackNo = oDS_PS_PP095L.GetValue("U_PackNo", i).ToString().Trim();
                    if (i == (oMat01.VisualRowCount - 1))
                    {
                        PackNo2 = "123";
                    }
                    else
                    {
                        PackNo2 = oDS_PS_PP095L.GetValue("U_PackNo", i + 1).ToString().Trim();
                    }
                    if (PackNo != PackNo2)
                    {
                        sQry = " EXEC [PS_PP095_13] ";
                        sQry = sQry + "'" + PackNo + "'";

                        oRecordSet01.DoQuery(sQry);
                        result = oRecordSet01.Fields.Item("RESULT").Value;

                        S_Weight += result;
                    }
                }
                sQry = "SELECT U_MINRANGE FROM [@PSH_MULTI_PALLET] WHERE CODE ='STANDARD'";

                oRecordSet01.DoQuery(sQry);

                result2 = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value);

                maxrange = S_Weight + result2;
                minrange = S_Weight - result2;
                
                if (Convert.ToDouble(TCHECK) == 1)
                {
                    // 최대값과 최소값 비교하여 참인지 거짓인지 비교
                    if (T_Weight > maxrange || T_Weight < minrange)
                    {
                        errMessage = "";
                    }
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_PP095_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
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
        /// PS_PP095_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP095_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)  ////행추가여부
                {
                    oDS_PS_PP095L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP095L.Offset = oRow;
                oDS_PS_PP095L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP095_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP095'", "");
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP095_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;
            int i;
            string OrdNum;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.VALUE))
                {
                    errMessage = "거래처코드는 필수입니다.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.VALUE))
                {
                    errMessage = "전기일은 필수입니다.";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.VALUE))
                    {
                        errMessage = "LOT-NO는 필수입니다.";
                        throw new Exception();
                    }
                    else
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            OrdNum = oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.VALUE;
                            Query01 = "select Cnt = Count(U_OrdNum) from [@PS_PP095H] a inner Join [@PS_PP095L] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
                            Query01 = Query01 + " Where b.U_OrdNum = '" + OrdNum + "'";
                            oRecordSet01.DoQuery(Query01);

                            if (oRecordSet01.Fields.Item("Cnt").Value > 0)
                            {
                                errMessage = OrdNum + " : 등록된 Lot번호입니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                oDS_PS_PP095L.RemoveRecord(oDS_PS_PP095L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP095_FormClear();
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
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP095_Print_Report01 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP095_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_PP095] 거래명세서출력";
                if (oForm.Items.Item("BPLId").Specific.Selected.VALUE == "2")
                {
                    ReportName = "PS_PP095_05.rpt"; //부산사업장용
                }
                else
                {
                    ReportName = "PS_PP095_01.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.VALUE)); //사업장

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP095_Print_Report02 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP095_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_PP095_02] 납품명세서출력";
                ReportName = "PS_PP095_02.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.VALUE)); //사업장

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP095_Print_Report03 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP095_Print_Report03()
        {
            string WinTitle;
            string ReportName;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_PP095_03] PACKING라벨출력";
                ReportName = "PS_PP095_03.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.VALUE)); //사업장

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_PP095_Print_Report04
        /// </summary>
        [STAThread]
        private void PS_PP095_Print_Report04()
        {
            int i;
            string WinTitle;
            string ReportName;
            string sQry01;
            string sQry02;
            string FilePath; 
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                FilePath = "\\\\191.1.1.220\\B1_SHR\\QRCODE_PACKING";

                sQry01 = "EXEC PS_PP095_99 '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'";
                oRecordSet01.DoQuery(sQry01);

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    ZXing.BarcodeWriter barcodeWriter = new ZXing.BarcodeWriter();
                    barcodeWriter.Format = ZXing.BarcodeFormat.QR_CODE;
                    barcodeWriter.Options.Margin = 0; // 이미지여백
                    barcodeWriter.Options.Width = 228; // 이미지 넓이
                    barcodeWriter.Options.Height = 228; // 이미지 높이
                    barcodeWriter.Write(oRecordSet01.Fields.Item(1).Value).Save(FilePath + "\\" + oRecordSet01.Fields.Item(0).Value + ".jpg", ImageFormat.Jpeg);

                    sQry02 = "Insert Into ZPS_PP095_QRCODE(DocEntry, PackNo) ";
                    sQry02 = sQry02 + " Select '" + oForm.Items.Item("DocEntry").Specific.VALUE + "','" + oRecordSet01.Fields.Item(0).Value + "'";
                    oRecordSet02.DoQuery(sQry02);

                    sQry02 = "Update ZPS_PP095_QRCODE Set QRImg = (Select Bulkcolumn From OPENROWSET(BULK N'" + FilePath + "\\" + oRecordSet01.Fields.Item(0).Value + ".jpg', SINGLE_BLOB) As QRImg)";
                    sQry02 = sQry02 + " Where DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "' And PackNo = '" + oRecordSet01.Fields.Item(0).Value + "'";
                    oRecordSet02.DoQuery(sQry02);

                    oRecordSet01.MoveNext();
                }
                WinTitle = "[PS_PP095_03_QRCODE] PACKING라벨(QRCODE)출력";
                ReportName = "PS_PP095_03_QRCODE.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.VALUE)); //사업장

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            string YM;
            //object ChildForm01 = null;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP095_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP095_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_PP095_Weight_Check() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_PP095_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_PP095_Print_Report02);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    if (pVal.ItemUID == "Button03")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                            if (PS_PP095_Check_QRCode_PrintYN() == "N")
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PS_PP095_Print_Report03);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start(); //일반 패킹리스트 출력
                            }
                            else if (PS_PP095_Check_QRCode_PrintYN() == "Y")
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PS_PP095_Print_Report04);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start(); //QR코드 패킹리스트 출력
                            }
                        }
                    }
                    if (pVal.ItemUID == "Button04")
                    {
                        YM = codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.VALUE, 1, 4) + "-" + codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.VALUE, 5, 2);

                        PS_QM041 ChildForm01 = new PS_QM041();
                        ChildForm01.LoadForm(YM, oForm.Items.Item("DocEntry").Specific.VALUE);
                    }
                    if (pVal.ItemUID == "Button05")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            //TempForm01 = new PS_MM004();
                            //TempForm01.LoadForm("PS_PP095", oForm.Items.Item("DocEntry").Specific.VALUE);
                            //BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Button06")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                            //PS_PP095_Print_Report04();
                        }
                    }
                    if (pVal.ItemUID == "Button07")
                    {
                        YM = codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.VALUE, 1, 4) + "-" + codeHelpClass.Mid(oForm.Items.Item("DocDate").Specific.VALUE, 5, 2);

                        PS_QM620 ChildForm01 = new PS_QM620(); 
                        ChildForm01.LoadForm(YM, oForm.Items.Item("DocEntry").Specific.VALUE);
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
                                PS_PP095_FormItemEnabled();
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP095_FormItemEnabled();
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
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.VALUE))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "OrdNum")
                            {
                                //ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, oForm.Items.Item("BPLId").Specific.VALUE);
                                //BubbleEvent = false;
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
                            oMat01.SelectRow(pVal.Row, true, false);
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
            int i;
            string Query01;
            double BoxWeight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "OrdNum")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.VALUE))
                                {
                                    throw new Exception();
                                }
                                Query01 = "SELECT PackNo   = a.U_PackNo, ";
                                Query01 = Query01 + " ItemCode = b.U_ItemCode, ";
                                Query01 = Query01 + " ItemName = b.U_ItemName, ";
                                Query01 = Query01 + " Weight   = b.U_Weight, ";
                                Query01 = Query01 + " LWeight  =  isnull((select max(boxwgt) as result from z_packlist where ordnum  in (select u_ordnum from [@ps_pp095l] where u_packno = a.U_PackNo)),0), ";
                                Query01 = Query01 + " ProDate = B.U_ProDate ";
                                Query01 = Query01 + " FROM [@PS_PP090H] a INNER JOIN [@PS_PP090L] b ON a.DocEntry = b.DocEntry AND a.CanCeled = 'N' ";
                                Query01 = Query01 + " WHERE ";
                                Query01 = Query01 + " a.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.VALUE + "'";
                                Query01 = Query01 + " AND b.U_LotNo = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.VALUE + "'";

                                oRecordSet01.DoQuery(Query01);

                                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                                {
                                    oDS_PS_PP095L.SetValue("U_OrdNum", pVal.Row - 1, oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.VALUE);
                                    oDS_PS_PP095L.SetValue("U_PackNo", pVal.Row - 1, oRecordSet01.Fields.Item("PackNo").Value);
                                    oDS_PS_PP095L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP095L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP095L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(oRecordSet01.Fields.Item("Weight").Value));
                                    oDS_PS_PP095L.SetValue("U_ProDate", pVal.Row - 1, oRecordSet01.Fields.Item("ProDate").Value.ToString("yyyyMMdd"));
                                    oRecordSet01.MoveNext();
                                }

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP095L.GetValue("U_OrdNum", pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP095_AddMatrixRow(pVal.Row, false);
                                }

                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oForm.Update();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        }
                        else
                        {
                            if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP095H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.VALUE);
                                oDS_PS_PP095H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'", 0, 1));
                            }
                            if (pVal.ItemUID == "CardCode")
                            {
                                oDS_PS_PP095H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.VALUE);
                                oDS_PS_PP095H.SetValue("U_CardName", 0, dataHelpClass.GetValue("select cardname from ocrd where cardcode = '" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'", 0, 1));
                            }
                            if (pVal.ItemUID == "TWeight")
                            {
                                if (oForm.Items.Item("TWeight").Specific.VALUE > 0)
                                {
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        BoxWeight = System.Math.Round((Convert.ToDouble(oForm.Items.Item("TWeight").Specific.VALUE) - Convert.ToDouble(oForm.Items.Item("S_Weight").Specific.VALUE)) / Convert.ToDouble(oForm.Items.Item("BoxCnt").Specific.VALUE), 1);
                                        oForm.Items.Item("Comments").Specific.VALUE = Convert.ToString(Convert.ToInt16(oForm.Items.Item("BoxCnt").Specific.VALUE)) + "EA X " + Convert.ToString(BoxWeight) + "Kg = " + Convert.ToString(Convert.ToDouble(oForm.Items.Item("TWeight").Specific.VALUE) - Convert.ToDouble(oForm.Items.Item("S_Weight").Specific.VALUE)) + " Kg";
                                    }
                                }
                                else
                                {
                                    oForm.Items.Item("Comments").Specific.VALUE = "";
                                }
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Update();
                oForm.Freeze(false);
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
            int i;
            int BoxCnt;
            double result;
            double SS_Weight = 0;
            string sQry; 
            string PackNo;
            string PackNo2;
            double S_Weight = 0;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP095_FormItemEnabled();
                    PS_PP095_AddMatrixRow(oMat01.VisualRowCount, false);

                    PackNo = oMat01.Columns.Item("PackNo").Cells.Item(1).Specific.VALUE;
                    BoxCnt = 0;
                    PackNo2 = "0";
                    if (oMat01.VisualRowCount <= 100)
                    {
                        for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            PackNo = oDS_PS_PP095L.GetValue("U_PackNo", i).ToString().Trim();
                            S_Weight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.VALUE);

                            if (i == (oMat01.VisualRowCount - 1))
                            {
                                PackNo2 = "123";
                            }
                            else
                            {
                                PackNo2 = oDS_PS_PP095L.GetValue("U_PackNo", i + 1).ToString().Trim();
                            }
                            if (PackNo != PackNo2)
                            {

                                sQry = " EXEC [PS_PP095_13] ";
                                sQry = sQry + "'" + PackNo + "'";

                                oRecordSet01.DoQuery(sQry);

                                result = Convert.ToDouble(oRecordSet01.Fields.Item("RESULT").Value);

                                SS_Weight += result;
                                BoxCnt += 1;
                            }
                        }
                        oForm.Items.Item("BoxCnt").Specific.VALUE = BoxCnt;
                        oForm.Items.Item("S_Weight").Specific.VALUE = S_Weight;
                        oForm.Items.Item("SS_Weight").Specific.VALUE = SS_Weight;
                    }
                    else
                    {
                        oForm.Items.Item("BoxCnt").Specific.VALUE = 0;
                        oForm.Items.Item("S_Weight").Specific.VALUE = 0;
                    }
                }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP095H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP095L);
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "CardCode" || pVal.ItemUID == "CardName")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP095H", "U_CardCode,U_CardName", "", 0, "", "", "");
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
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
                    oDS_PS_PP095L.RemoveRecord(oDS_PS_PP095L.Size - 1);
                    oMat01.LoadFromDataSource();
                    if (oMat01.RowCount == 0)
                    {
                        PS_PP095_AddMatrixRow(0, false);
                    }
                    oForm.Update();
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
                            PS_PP095_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP095_AddMatrixRow(0, true);
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
