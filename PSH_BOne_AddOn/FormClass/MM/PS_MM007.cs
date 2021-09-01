using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using SAP.Middleware.Connector;
namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 품목마스터 생성요청
    /// </summary>
    internal class PS_MM007 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM007H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM007L; //등록라인

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM007.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM007_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM007");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                if (!string.IsNullOrEmpty(oFromDocEntry01))
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                oForm.Freeze(true);
                PS_MM007_CreateItems();
                PS_MM007_ComboBox_Setting();

                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), true); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1284"), true); // 취소
                oForm.EnableMenu(("1293"), true); // 행삭제

                if (!string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_MM007_FormItemEnabled();
                    oForm.Items.Item("DocNum").Specific.Value = oFromDocEntry01;
                    oForm.Items.Item("CntcCode").Specific.Value = "";
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PS_MM007_Initialization();
                }
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
        private void PS_MM007_CreateItems()
        {
            try
            {
                oDS_PS_MM007H = oForm.DataSources.DBDataSources.Item("@PS_MM007H");
                oDS_PS_MM007L = oForm.DataSources.DBDataSources.Item("@PS_MM007L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oDS_PS_MM007H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM007_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oForm.Items.Item("ItmsGrp").Specific.ValidValues.Add("105", "[5]저장품");
                oDS_PS_MM007H.SetValue("U_ItmsGrp", 0, "105");

                oForm.Items.Item("RFCAdms").Specific.ValidValues.Add("N", "미승인");
                oForm.Items.Item("RFCAdms").Specific.ValidValues.Add("Y", "승인");
                oDS_PS_MM007H.SetValue("U_RFCAdms", 0, "N");
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
        /// Initialization
        /// </summary>
        private void PS_MM007_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM007_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM007H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM007H.GetValue("U_CntcCode", 0)))
                {
                    errMessage = "담당자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM007H.GetValue("U_OffcTel", 0)))
                {
                    errMessage = "전화번호는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM007H.GetValue("U_ItmsGrp", 0)))
                {
                    errMessage = "품목그룹은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM007H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수사항입니다. 확인하세요.";
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM007_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            int i; 
            string sQry;
            string ReqNo;
            string Seq;
            string DocDate;
            string errMessage = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_ReqNo", 0)))
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                sQry = "Select ReqNo = Max(U_ReqNo) From [@PS_MM007L] Where Left(U_ReqNo,8) = '" + DocDate + "'";
                oRecordSet01.DoQuery(sQry);

                if (string.IsNullOrEmpty(oRecordSet01.Fields.Item("ReqNo").Value.ToString().Trim()))
                {
                    ReqNo = DocDate + "000";
                    Seq = "000";
                }
                else
                {
                    ReqNo = oRecordSet01.Fields.Item("ReqNo").Value.ToString().Trim();
                    Seq = codeHelpClass.Right(oRecordSet01.Fields.Item("ReqNo").Value.ToString().Trim(), 3);
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oMat01.Columns.Item("ReqDate").Cells.Item(i).Specific.Value = DocDate;
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ReqNo").Cells.Item(i).Specific.Value.ToString().Trim()))
                    {
                        oMat01.Columns.Item("ReqNo").Cells.Item(i).Specific.Value = codeHelpClass.Left(ReqNo, 8) + codeHelpClass.Right("000" + (Convert.ToString(Convert.ToInt32(Seq) + i)), 3);
                    }
                }

                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_ReqNo", i)))
                    {
                        errMessage = "청구번호는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_TeamName", i)))
                    {
                        errMessage = "팀명은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_MATNR", i)))
                    {
                        errMessage = "상품번호 7자리는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_MAKTX", i)))
                    {
                        errMessage = "자재내역은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_NORMT", i)))
                    {
                        errMessage = "메이커는 필수사항입니다. 없을시'국산'이라고 입력바랍니다.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_MEINS", i)))
                    {
                        errMessage = "기본단위는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (oDS_PS_MM007L.GetValue("U_MEINS", i).ToString().Trim().Length > 2)
                    {
                        errMessage = "기본단위는 두 글자로 입력하셔야 합니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_MATKL", i)))
                    {
                        errMessage = "상품범주는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();
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
        /// Delete_EmptyRow
        /// </summary>
        private void PS_MM007_Delete_EmptyRow()
        {
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_ReqNo", i).ToString().Trim()))
                    {
                        oDS_PS_MM007L.RemoveRecord(i);
                    }
                }
                oMat01.LoadFromDataSource();
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
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM007_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("ItmsGrp").Enabled = true;
                    oForm.Items.Item("RFCAdms").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oDS_PS_MM007H.SetValue("U_ItmsGrp", 0, "105"); //품목그룹
                    oDS_PS_MM007H.SetValue("U_RFCAdms", 0, "N"); //승인여부 
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("ItmsGrp").Enabled = true;
                    oForm.Items.Item("OffcTel").Enabled = true;
                    oForm.Items.Item("RFCAdms").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;

                    if (oDS_PS_MM007H.GetValue("U_RFCAdms", 0).ToString().Trim() == "Y" || oDS_PS_MM007H.GetValue("Canceled", 0).ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("CntcCode").Enabled = false;
                        oForm.Items.Item("ItmsGrp").Enabled = false;
                        oForm.Items.Item("OffcTel").Enabled = false;
                        oForm.Items.Item("RFCAdms").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("BPLId").Enabled = true;
                        oForm.Items.Item("CntcCode").Enabled = true;
                        oForm.Items.Item("ItmsGrp").Enabled = true;
                        oForm.Items.Item("OffcTel").Enabled = true;
                        oForm.Items.Item("RFCAdms").Enabled = true;
                        oForm.Items.Item("DocDate").Enabled = true;
                        oForm.Items.Item("Mat01").Enabled = true;
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
        /// PS_MM007_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM007_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM007L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM007L.Offset = oRow;
                oDS_PS_MM007L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_MM007_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM007'", "");
                if (Convert.ToDouble(DocNum) == 0)
                {
                    oForm.Items.Item("DocNum").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM007_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                switch (oUID)
                {
                    case "CntcCode":
                        sQry = "Select U_FullName, U_OffcTel, U_eMail From [@PH_PY001A] Where Code = '" + oDS_PS_MM007H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oDS_PS_MM007H.SetValue("U_CntcName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        oDS_PS_MM007H.SetValue("U_OffcTel", 0, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                        oDS_PS_MM007H.SetValue("U_EMail", 0, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                        break;

                    case "Mat01":
                        if (oCol == "TeamCode") //부서
                        {
                            if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("TeamCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                oMat01.FlushToDataSource();
                                PS_MM007_AddMatrixRow(oMat01.RowCount, false);
                                oMat01.Columns.Item("TeamCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            sQry = "Select U_CodeNm From [@PS_HR200L] Where Code = '1' and U_Code = '" + oMat01.Columns.Item("TeamCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("TeamName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CodeNm").Value.ToString().Trim();
                        }
                        else if (oCol == "RspCode")
                        {
                            sQry = "Select U_CodeNm From [@PS_HR200L] Where Code = '2' and U_Code = '" + oMat01.Columns.Item("RspCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("RspName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CodeNm").Value.ToString().Trim();
                           
                        }
                        else if (oCol == "MATKL")  //품목분류
                        {
                            sQry = "Select CodeName = U_CodeName From [@PSH_ITMMSORT] Where U_Code = '" + oMat01.Columns.Item("MATKL").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("MATKLNM").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("CodeName").Value.ToString().Trim();
                        }
                        break;
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_RFC_Sender
        /// </summary>
        /// <param name="BPLID"></param>
        /// <param name="RequestNo"></param>
        /// <param name="CntcName"></param>
        /// <param name="TeamName"></param>
        /// <param name="ReqDate"></param>
        /// <param name="OffcTel"></param>
        /// <param name="EMail"></param>
        /// <param name="MATNR"></param>
        /// <param name="MAKTX"></param>
        /// <param name="MTART"></param>
        /// <param name="MEINS"></param>
        /// <param name="MATKL"></param>
        /// <param name="ZUSE"></param>
        /// <param name="NORMT"></param>
        /// <param name="i"></param>
        /// <param name="LastRow"></param>
        /// <returns></returns>
        private string PS_RFC_Sender(string BPLID, string RequestNo, string CntcName, string TeamName, string ReqDate, string OffcTel, string EMail, string MATNR, string MAKTX, string MTART, string MEINS, string MATKL, string ZUSE, string NORMT, int i, int LastRow)
        {
            string returnValue = string.Empty;
            string WERKS = string.Empty;
            string errMessage = string.Empty;
            string Client; //클라이언트(운영용:210, 테스트용:710)
            string ServerIP;
            string errCode = string.Empty;
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;
            IRfcFunction oFunction = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //서버IP (운영용:192.1.11.3, 테스트용:192.1.11.7)
                //Real
                Client = "210";
                ServerIP = "192.1.11.3";

                ////test
                //Client = "810";
                //ServerIP = "192.1.11.7";

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errMessage = "풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.";
                    throw new Exception();
                }

                oFunction = rfcRep.CreateFunction("ZMM_SUB_REQMAT");
                
                switch (BPLID)
                {
                    case "1":
                        WERKS = "9200";
                        break;
                    case "2":
                        WERKS = "9300";
                        break;
                    case "3":
                        WERKS = "9500";
                        break;
                    case "5":
                        WERKS = "9600";
                        break;
                }

                oFunction.SetValue("I_WERKS", WERKS); //플랜트 홀딩스 창원 9200, 홀딩스 부산 9300, 포장사업팀 9500, 포장온산 9600
                oFunction.SetValue("I_REQNO", RequestNo); //요청문서번호
                oFunction.SetValue("I_MATNR", MATNR); //상품번호(7자리)
                oFunction.SetValue("I_MAKTX", MAKTX); //상품내역
                oFunction.SetValue("I_WRKST", ""); //기본자재
                oFunction.SetValue("I_USRNB", CntcName); //요청생성인
                oFunction.SetValue("I_DEPRT", TeamName); //부서 담당
                oFunction.SetValue("I_ERSDA", ReqDate); //요청일자
                oFunction.SetValue("I_TELNO", OffcTel); //전화번호
                oFunction.SetValue("I_EMAIL", EMail); //이메일
                oFunction.SetValue("I_MTART", MTART); //자재유형
                oFunction.SetValue("I_MEINS", MEINS); //기본단위
                oFunction.SetValue("I_MATKL", MATKL); //상품범주
                oFunction.SetValue("I_NORMT", NORMT); //(메이커)MAKER
                oFunction.SetValue("I_ZUSE", ZUSE); //용도
                oFunction.SetValue("I_ZSUPP", ""); //첨부물표시
                oFunction.SetValue("I_FILENAME", ""); //첨부파일이름

                errCode = "E2"; // 아래 invoke 오류 체크를 위한 변수대입
                oFunction.Invoke(rfcDest); //Function 실행
                errCode = string.Empty;// 이상 없을 경우 초기화
                if (string.IsNullOrEmpty(oFunction.GetValue("E_MESSAGE").ToString()))
                {
                    oDS_PS_MM007L.SetValue("U_MESSAGE", i, "");
                    oDS_PS_MM007L.SetValue("U_MESSAGE", i, oFunction.GetValue("E_MESSAGE").ToString().Trim()); //에러메세지
                    oDS_PS_MM007L.SetValue("U_TransYN", i, oFunction.GetValue("E_TransYN").ToString().Trim()); //전송접수여부
                    errMessage = oFunction.GetValue("E_MESSAGE").ToString().Trim();
                }
                else
                {
                    returnValue = oFunction.GetValue("E_TransYN").ToString().Trim();
                    oDS_PS_MM007L.SetValue("U_TransYN", i, oFunction.GetValue("E_TransYN").ToString().Trim());
                    oDS_PS_MM007L.SetValue("U_MESSAGE", i, oFunction.GetValue("E_MESSAGE").ToString().Trim());
                }
            }
            catch (Exception ex)
            {
                if (errCode == "E2")
                {
                    PSH_Globals.SBO_Application.MessageBox("통합구매(R/3) 함수(ZMM_SUB_REQMAT)호출중 오류발생");
                }
                else if (errMessage != string.Empty)
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
            int i;
            string BPLID;
            string DueDate;
            string CntcName;
            string OffcTel;
            string TeamName;
            string RequestDate;
            string RequestNo;
            string RFC_Sender;
            string EMail;
            string MATNR;
            string MAKTX;
            string MTART;
            string MEINS;
            string MATKL;
            string ZUSE;
            string NORMT;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM007_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM007_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //신규코드 생성요청
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && oDS_PS_MM007H.GetValue("U_RFCAdms", 0).ToString().Trim() == "Y")
                            {
                                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("통합구매 신규코드생성 요청!", oMat01.VisualRowCount - 2 + 1, false);
                                oMat01.FlushToDataSource();

                                DueDate = oDS_PS_MM007H.GetValue("U_DocDate", 0).ToString().Trim();
                                BPLID = oDS_PS_MM007H.GetValue("U_BPLId", 0).ToString().Trim();
                                CntcName = oDS_PS_MM007H.GetValue("U_CntcName", 0).ToString().Trim();
                                OffcTel = oDS_PS_MM007H.GetValue("U_OffcTel", 0).ToString().Trim();
                                EMail = oDS_PS_MM007H.GetValue("U_EMail", 0).ToString().Trim();

                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    if (oDS_PS_MM007L.GetValue("U_TransYN", i).ToString().Trim() == "N" || string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_TransYN", i).ToString().Trim()))
                                    {
                                        RequestDate = oDS_PS_MM007L.GetValue("U_ReqDate", 0).ToString().Trim(); //요청일자
                                        RequestNo = oDS_PS_MM007L.GetValue("U_ReqNo", i).ToString().Trim(); //요청번호

                                        if (string.IsNullOrEmpty(oDS_PS_MM007L.GetValue("U_RspName", i).ToString().Trim()))
                                        {
                                            TeamName = oDS_PS_MM007L.GetValue("U_TeamName", i).ToString().Trim();
                                        }
                                        else
                                        {
                                            TeamName = oDS_PS_MM007L.GetValue("U_TeamName", i).ToString().Trim() + " / " + oDS_PS_MM007L.GetValue("U_RspName", i).ToString().Trim(); //부서/담당명
                                        }
                                        MATNR = oDS_PS_MM007L.GetValue("U_MATNR", i).ToString().Trim(); //상품번호
                                        MAKTX = oDS_PS_MM007L.GetValue("U_MAKTX", i).ToString().Trim(); //상품내역
                                        MTART = oDS_PS_MM007L.GetValue("U_MTART", i).ToString().Trim(); //상품유형
                                        MEINS = oDS_PS_MM007L.GetValue("U_MEINS", i).ToString().Trim(); //기본단위
                                        MATKL = oDS_PS_MM007L.GetValue("U_MATKL", i).ToString().Trim(); //상품범주
                                        ZUSE = oDS_PS_MM007L.GetValue("U_ZUSE", i).ToString().Trim(); //용도
                                        NORMT = oDS_PS_MM007L.GetValue("U_NORMT", i).ToString().Trim(); //메이커(MAKER)

                                        //사업장, 요청번호, 팀/담당명, 요청일자, 전화번호, 이메일, 상품번호(7자리), 상품내역, 상품유형, 기본단위, 상품범주, 메이커,
                                        RFC_Sender = PS_RFC_Sender(BPLID, RequestNo, CntcName, TeamName, RequestDate, OffcTel, EMail, MATNR, MAKTX, MTART, MEINS, MATKL, ZUSE, NORMT, i, oMat01.VisualRowCount - 2);

                                        if (!string.IsNullOrEmpty(RFC_Sender))
                                        {
                                            oDS_PS_MM007L.SetValue("U_TransYN", i, RFC_Sender);
                                        }
                                        ProgressBar01.Value += 1;
                                        ProgressBar01.Text = Convert.ToString(ProgressBar01.Value) + "/" + Convert.ToString(oMat01.VisualRowCount - 2 + 1) + "건 처리중...!";
                                    }
                                }
                                oMat01.LoadFromDataSource();
                            }
                            PS_MM007_Delete_EmptyRow();
                        }
                        oMat01.AutoResizeColumns();
                    }
                    
                    if (pVal.ItemUID == "Btn_01")
                    {
                        System.Diagnostics.Process ps = new System.Diagnostics.Process();
                        ps.StartInfo.FileName = "RFC_INTERFACE_OITM.exe";
                        ps.StartInfo.WorkingDirectory = "\\\\191.1.1.221\\ERP_Project\\9999. Add_On\\인터페이스05(품목생성요청)\\";
                        ps.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                        ps.Start();
                        ps.WaitForExit(1000);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                if(ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "TeamCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("TeamCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ColUID == "RspCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ColUID == "MATKL")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("MATKL").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
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
                    if (pVal.ItemUID == "ItmsGrp" || pVal.ItemUID == "BPLId")
                    {
                        oMat01.Clear();
                        oDS_PS_MM007L.Clear();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM007_AddMatrixRow(0, false);
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM007_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "TeamCode" || pVal.ColUID == "RspCode")
                            {
                                PS_MM007_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "MATNR")
                            {
                                PS_MM007_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "MATKL")
                            {
                                PS_MM007_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);

                            }
                        }
                        oMat01.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM007H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM007L);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oForm.Items.Item("RFCAdms").Specific.Value == "Y")
                            {
                                dataHelpClass.MDC_GF_Message("통합구매 코드생성한 자료입니다. 취소할 수 없습니다.", "E");
                                BubbleEvent = false;
                                return;
                            }
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
                                PS_MM007_FormItemEnabled();
                                oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                break;
                            case "1286": //닫기
                                break;
                            case "1293": //행삭제
                                if (oMat01.RowCount != oMat01.VisualRowCount)
                                {
                                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                    {
                                        oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                    }

                                    oMat01.FlushToDataSource();
                                    oDS_PS_MM007L.RemoveRecord(oDS_PS_MM007L.Size - 1); // Mat01에 마지막라인(빈라인) 삭제
                                    oMat01.Clear();
                                    oMat01.LoadFromDataSource();
                                    oMat01.AutoResizeColumns();
                                }
                                break;
                            case "1281": //찾기
                                PS_MM007_FormItemEnabled();
                                oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            
                                if (dataHelpClass.User_SuperUserYN() == "Y") //슈퍼유저인경우는 사번 추가 하지 않음(2018.04.02 송명규)
                                {
                                    oForm.Items.Item("CntcCode").Specific.Value = "";
                                }
                                else
                                {
                                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                                }
                                oMat01.AutoResizeColumns();
                                break;
                            case "1282": //추가
                                PS_MM007_Initialization();
                                PS_MM007_FormItemEnabled();
                                PS_MM007_FormClear();
                                oDS_PS_MM007H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                                PS_MM007_AddMatrixRow(0, true);
                                oMat01.AutoResizeColumns();
                                break;
                            case "1288": //레코드이동(최초)
                            case "1289": //레코드이동(이전)
                            case "1290": //레코드이동(다음)
                            case "1291": //레코드이동(최종)
                                PS_MM007_FormItemEnabled();
                                if (oMat01.VisualRowCount > 0)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("ReqNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                    {
                                        if (oDS_PS_MM007H.GetValue("Status", 0) == "O")
                                        {
                                            PS_MM007_AddMatrixRow(oMat01.RowCount, false);
                                        }
                                    }
                                }
                                oMat01.AutoResizeColumns();
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
