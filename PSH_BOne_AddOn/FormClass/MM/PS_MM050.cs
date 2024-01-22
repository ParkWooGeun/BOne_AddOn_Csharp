using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 가입고
    /// </summary>
    internal class PS_MM050 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_MM050H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM050L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_TEMPTABLE;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM050.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM050_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM050");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_MM050_CreateItems();
                PS_MM050_ComboBox_Setting();
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PS_MM050_Initialization();
                }
                PS_MM050_FormItemEnabled();
                PS_MM050_FormClear();
                PS_MM050_FormResize();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1286", true); //닫기
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    oForm.Items.Item("DocNum").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM050_CreateItems()
        {
            try
            {
                oDS_PS_TEMPTABLE = oForm.DataSources.DBDataSources.Item("@PS_TEMPTABLE");
                oDS_PS_MM050H = oForm.DataSources.DBDataSources.Item("@PS_MM050H");
                oDS_PS_MM050L = oForm.DataSources.DBDataSources.Item("@PS_MM050L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oForm.DataSources.UserDataSources.Add("DueDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DueDateFr").Specific.DataBind.SetBound(true, "", "DueDateFr");

                oForm.DataSources.UserDataSources.Add("DueDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DueDateTo").Specific.DataBind.SetBound(true, "", "DueDateTo");
                oDS_PS_MM050H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                oForm.DataSources.UserDataSources.Add("DocTotal", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("DocTotal").Specific.DataBind.SetBound(true, "", "DocTotal");

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM050_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                //품목구분
                sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oMat01.Columns.Item("ItemGpCd").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oMat02.Columns.Item("ItemGpCd").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                //품의형태
                sQry = "SELECT Code, Name From [@PSH_RETYPE]";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oForm.Items.Item("POType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("POType").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oMat02.Columns.Item("POType").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                //품의상태
                oForm.Items.Item("POStatus").Specific.ValidValues.Add("Y", "승인");
                oForm.Items.Item("POStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품질검수여부
                oMat02.Columns.Item("QEYesNo").ValidValues.Add("Y", "Yes");
                oMat02.Columns.Item("QEYesNo").ValidValues.Add("N", "No");
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
        /// Initialization
        /// </summary>
        private void PS_MM050_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();//아이디별 사번 세팅
                oForm.Items.Item("POStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
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
        private bool PS_MM050_HeaderSpaceLineDel()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_CardCode", 0)))
                {
                    errMessage = "거래처는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_CntcCode", 0)))
                {
                    errMessage = "담당자는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_POType", 0)))
                {
                    errMessage = "품의형태는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_POStatus", 0)))
                {
                    errMessage = "품의상태은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM050H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수입력사항입니다. 확인하세요.";
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
        private bool PS_MM050_MatrixSpaceLineDel()
        {
            bool ReturnValue = false;
            int i;
            int ErrRowCount = 0;
            string sQry;
            int errCode = 0;
            string errMessage = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oMat02.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                else
                {
                    for (i = 1; i <= oMat02.VisualRowCount; i++)
                    {
                        sQry = "exec PS_MM050_02 '" + oMat02.Columns.Item("PODocNum").Cells.Item(i).Specific.Value + "','" + oMat02.Columns.Item("POLinNum").Cells.Item(i).Specific.Value + "','" + oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "'," + Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value) + "";
                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item(0).Value == "E")
                        {
                            if (PSH_Globals.SBO_Application.MessageBox(oRecordSet.Fields.Item(1).Value + "계속 진행하시겠습니까?", 2, "Yes", "No") == 2)
                            {
                                errCode = 1;
                                throw new Exception();
                            }
                        }

                        if (string.IsNullOrEmpty(oMat02.Columns.Item("Qty").Cells.Item(i).Specific.Value))
                        {
                            errMessage = Convert.ToString(ErrRowCount) + "행의 수량에 값이 없습니다. 확인바랍니다.";
                            ErrRowCount = i;
                            throw new Exception();
                        }
                        else if (PS_MM050_CheckDate(oMat02.Columns.Item("PODocNum").Cells.Item(i).Specific.Value) == false)
                        {
                            errMessage = ErrRowCount + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(ErrRowCount + 1).Specific.Value + "]의 가입고일은 구매품의일과 같거나 늦어야합니다. 확인하십시오. 해당 가입고는 전체가 등록되지 않습니다.";
                            ErrRowCount = i;
                            throw new Exception();
                        }
                    }
                }
                oMat01.LoadFromDataSource();
                ReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("취소처리되었습니다.");
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
            return ReturnValue;
        }

        /// <summary>
        /// PS_MM050_Display_MatrixData
        /// </summary>
        /// <returns></returns>
        private bool PS_MM050_Display_MatrixData()
        {
            bool returnValue = false;
            int sCnt;
            string sQry;
            string DueDateFr;
            string POType;
            string BPLId;
            string CardCode;
            string CntcCode;
            string POStatus;
            string DueDateTo;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);
                oMat01.Clear();
                oMat02.Clear();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                POType = oForm.Items.Item("POType").Specific.Value.ToString().Trim();
                POStatus = oForm.Items.Item("POStatus").Specific.Value.ToString().Trim();
                DueDateFr = oForm.Items.Item("DueDateFr").Specific.Value.ToString().Trim();
                DueDateTo = oForm.Items.Item("DueDateTo").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(CardCode))
                {
                    CardCode = "%";
                }
                if (string.IsNullOrEmpty(BPLId))
                {
                    BPLId = "%";
                }
                if (string.IsNullOrEmpty(CntcCode))
                {
                    CntcCode = "%";
                }
                if (string.IsNullOrEmpty(POType))
                {
                    POType = "%";
                }
                if (string.IsNullOrEmpty(POStatus))
                {
                    POStatus = "%";
                }
                if (string.IsNullOrEmpty(DueDateFr))
                {
                    DueDateFr = "20100101";
                }
                if (string.IsNullOrEmpty(DueDateTo))
                {
                    DueDateTo = "20991231";
                }

                sQry = "EXEC [PS_MM050_01] '" + CardCode + "','" + BPLId + "','" + CntcCode + "','" + POType + "','" + POStatus + "','" + DueDateFr + "','" + DueDateTo + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oMat02.Clear();
                oDS_PS_TEMPTABLE.Clear();

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                sCnt = 0;
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oDS_PS_TEMPTABLE.InsertRecord(sCnt);
                        oDS_PS_TEMPTABLE.Offset = sCnt;
                        oDS_PS_TEMPTABLE.SetValue("U_iField01", sCnt, Convert.ToString(sCnt + 1));
                        oDS_PS_TEMPTABLE.SetValue("U_sField01", sCnt, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_iField02", sCnt, oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_iField06", sCnt, oRecordSet.Fields.Item("EBELN").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField03", sCnt, oRecordSet.Fields.Item(2).Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField04", sCnt, oRecordSet.Fields.Item(3).Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField18", sCnt, oRecordSet.Fields.Item("U_ItemGpCd").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField05", sCnt, oRecordSet.Fields.Item(4).Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_dField01", sCnt, oRecordSet.Fields.Item(5).Value.ToString("yyyyMMdd").Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_dField02", sCnt, oRecordSet.Fields.Item(6).Value.ToString("yyyyMMdd").Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_iField03", sCnt, oRecordSet.Fields.Item("U_Qty").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_qField01", sCnt, oRecordSet.Fields.Item("U_Weight").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_qField02", sCnt, oRecordSet.Fields.Item("U_UnWeight").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_aField01", sCnt, oRecordSet.Fields.Item("U_Price").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_aField02", sCnt, oRecordSet.Fields.Item("U_LinTotal").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField12", sCnt, oRecordSet.Fields.Item("U_WhsCode").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField13", sCnt, oRecordSet.Fields.Item("U_WhsName").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_iField04", sCnt, oRecordSet.Fields.Item("PendNum").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField15", sCnt, oRecordSet.Fields.Item("U_PODocNum").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField16", sCnt, oRecordSet.Fields.Item("VisOrder").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField23", sCnt, oRecordSet.Fields.Item("U_OutSize").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField24", sCnt, oRecordSet.Fields.Item("U_OutUnit").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField25", sCnt, oRecordSet.Fields.Item("U_Auto").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField22", sCnt, oRecordSet.Fields.Item("U_Comments").Value.ToString().Trim());
                        oDS_PS_TEMPTABLE.SetValue("U_sField10", sCnt, oRecordSet.Fields.Item("U_ReqCntc").Value.ToString().Trim()); //청구자
                        oDS_PS_TEMPTABLE.SetValue("U_sField07", sCnt, oRecordSet.Fields.Item("U_OrdNum").Value.ToString().Trim()); //작번
                        oDS_PS_TEMPTABLE.SetValue("U_sField08", sCnt, oRecordSet.Fields.Item("U_OrdSub1").Value.ToString().Trim()); //Sub작번1
                        oDS_PS_TEMPTABLE.SetValue("U_sField09", sCnt, oRecordSet.Fields.Item("U_OrdSub2").Value.ToString().Trim()); //Sub작번2
                        oDS_PS_TEMPTABLE.SetValue("U_sField17", sCnt, oRecordSet.Fields.Item("U_DocCur").Value.ToString().Trim()); //통화
                        oDS_PS_TEMPTABLE.SetValue("U_qField04", sCnt, oRecordSet.Fields.Item("U_DocRate").Value.ToString().Trim()); //환율
                        oDS_PS_TEMPTABLE.SetValue("U_qField03", sCnt, oRecordSet.Fields.Item("U_FCPrice").Value.ToString().Trim()); //외화단가
                        oDS_PS_TEMPTABLE.SetValue("U_aField03", sCnt, oRecordSet.Fields.Item("U_FCAmount").Value.ToString().Trim()); //외화금액
                        sCnt += 1;
                        oRecordSet.MoveNext();
                    }
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
            return returnValue;
        }

        /// <summary>
        /// 선행프로세스와 일자 비교
        /// </summary>
        /// <returns>true:선행프로세스보다 일자가 같거나 느릴 경우, false:선행프로세스보다 일자가 빠를 경우</returns>
        private bool PS_MM050_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string sQry;
            string baseEntry;
            string baseLine;
            string docType;
            string CurDocDate;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                baseEntry = pBaseEntry;
                baseLine = "";
                docType = "PS_MM050";
                CurDocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                sQry = "EXEC PS_Z_CHECK_DATE '";
                sQry += baseEntry + "','";
                sQry += baseLine + "','";
                sQry += docType + "','";
                sQry += CurDocDate + "'";

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item("ReturnValue").Value != "False")
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return returnValue;
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM050_Validate(string ValidateType)
        {
            bool ReturnValue = false;
            int i; ;
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value))
                        {
                        }
                        else
                        {
                            sQry = "  SELECT    COUNT(*) AS [Count],";
                            sQry += "           T0.DocEntry AS [DocEntry],";
                            sQry += "           T1.U_LineNum AS [LineNum]";
                            sQry += " FROM      [@PS_MM070H] AS T0";
                            sQry += "           INNER JOIN";
                            sQry += "           [@PS_MM070L] AS T1";
                            sQry += "               ON T0.DocEntry = T1.DocEntry";
                            sQry += " WHERE     T1.U_GADocLin = '" + oForm.Items.Item("DocNum").Specific.Value + "-" + oMat02.Columns.Item("LineID").Cells.Item(oLastColRow01).Specific.Value + "'";
                            sQry += "           AND T0.Status = 'O'";
                            sQry += "           AND T1.U_RealWt <> 0";
                            sQry += " GROUP BY  T0.DocEntry,";
                            sQry += "           T1.U_LineNum";

                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.Fields.Item("Count").Value > 0)
                            {
                                errMessage = "선택한 품목은 이미 검수입고 되었습니다(검수입고 문서번호 : " + oRecordSet.Fields.Item("DocEntry").Value + "-" + oRecordSet.Fields.Item("LineNum").Value + "). 해당 가입고문서를 행삭제할 수 없습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    for (i = 1; i < oMat02.VisualRowCount; i++)
                    {
                        sQry = " SELECT    COUNT(*) AS [Count],";
                        sQry += "           T0.DocEntry AS [DocEntry],";
                        sQry += "           T1.U_LineNum AS [LineNum]";
                        sQry += " FROM      [@PS_MM070H] AS T0";
                        sQry += "           INNER JOIN";
                        sQry += "           [@PS_MM070L] AS T1";
                        sQry += "               ON T0.DocEntry = T1.DocEntry";
                        sQry += " WHERE     T1.U_GADocLin = '" + oForm.Items.Item("DocNum").Specific.Value + "-" + oMat02.Columns.Item("LineID").Cells.Item(i).Specific.Value + "'";
                        sQry += "           AND T0.Status = 'O'";
                        sQry += "           AND T1.U_RealWt <> 0";
                        sQry += " GROUP BY  T0.DocEntry,";
                        sQry += "           T1.U_LineNum";
                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item("Count").Value > 0)
                        {
                            errMessage = i + "행 품목이 이미 검수입고 되었습니다(검수입고 문서번호 : " + oRecordSet.Fields.Item("DocEntry").Value + "-" + oRecordSet.Fields.Item("LineNum").Value + "). 해당 가입고문서를 취소할 수 없습니다.";
                            throw new Exception();
                        }
                    }
                }
                else if (ValidateType == "닫기")
                {
                    for (i = 1; i < oMat02.VisualRowCount; i++)
                    {
                        sQry = " SELECT    COUNT(*) AS [Count],";
                        sQry += "           T0.DocEntry AS [DocEntry],";
                        sQry += "           T1.U_LineNum AS [LineNum]";
                        sQry += " FROM      [@PS_MM070H] AS T0";
                        sQry += "           INNER JOIN";
                        sQry += "           [@PS_MM070L] AS T1";
                        sQry += "               ON T0.DocEntry = T1.DocEntry";
                        sQry += " WHERE     T1.U_GADocLin = '" + oForm.Items.Item("DocNum").Specific.Value + "-" + oMat02.Columns.Item("LineID").Cells.Item(i).Specific.Value + "'";
                        sQry += "           AND T0.Status = 'O'";
                        sQry += " GROUP BY  T0.DocEntry,";
                        sQry += "           T1.U_LineNum";

                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item("Count").Value > 0)
                        {
                            errMessage = i + "행 품목이 이미 검수입고 되었습니다(검수입고 문서번호 : " + oRecordSet.Fields.Item("DocEntry").Value + "-" + oRecordSet.Fields.Item("LineNum").Value + "). 해당 가입고문서를 닫기할 수 없습니다.";
                            throw new Exception();
                        }
                    }
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return ReturnValue;
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_MM050_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat01").Top = 82;
                oForm.Items.Item("Mat01").Left = 6;
                oForm.Items.Item("Mat01").Width = oForm.Width - 18;
                oForm.Items.Item("Mat01").Height = (oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height - oForm.Items.Item("Comments").Top)) / 2 - 10;

                oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat01").Height + oForm.Items.Item("Mat01").Top;
                oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat01").Left;
                oForm.Items.Item("Mat02").Width = oForm.Items.Item("Mat01").Width;
                oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height - 5;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM050_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("POType").Enabled = true;
                    oForm.Items.Item("POStatus").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oForm.Items.Item("Btn01").Visible= true;

                    oMat02.Columns.Item("Qty").Editable = true;
                    oMat02.Columns.Item("Weight").Editable = true;
                    oMat02.Columns.Item("RealWt").Editable = true;
                    oMat02.Columns.Item("LinTotal").Editable = true;
                    oMat02.Columns.Item("UnWeight").Editable = true;
                    oMat02.Columns.Item("WhsCode").Editable = true;
                    oMat02.Columns.Item("QEYesNo").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("POType").Enabled = true;
                    oForm.Items.Item("POStatus").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Btn01").Visible = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("POType").Enabled = false;
                    oForm.Items.Item("POStatus").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat02.Columns.Item("LinTotal").Editable = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oMat02.Columns.Item("Qty").Editable = true;
                    oMat02.Columns.Item("Weight").Editable = true;
                    oMat02.Columns.Item("RealWt").Editable = true;
                    oMat02.Columns.Item("UnWeight").Editable = true;
                    oMat02.Columns.Item("WhsCode").Editable = true;
                    oMat02.Columns.Item("QEYesNo").Editable = true;
                    oForm.Items.Item("Btn01").Visible = false;
                    if (oForm.Items.Item("Status").Specific.Value == "C")
                    {
                        oForm.Items.Item("Mat02").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Mat02").Enabled = true;
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
        /// Delete_EmptyRow
        /// </summary>
        private void PS_MM050_Delete_EmptyRow()
        {
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat02.FlushToDataSource();
                for (i = 0; i < oMat02.VisualRowCount; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM050L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        oDS_PS_MM050L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
                    }
                }
                oMat02.LoadFromDataSource();
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
        /// PS_MM050_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM050_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_TEMPTABLE.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_TEMPTABLE.Offset = oRow;
                oDS_PS_TEMPTABLE.SetValue("U_iField01", oRow, Convert.ToString(oRow + 1));
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
        /// PS_MM050_AddMatrixRow02
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM050_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM050L.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_MM050L.Offset = oRow;
                oDS_PS_MM050L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
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
        private void PS_MM050_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM050'", "");
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
        private void PS_MM050_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i;
            int sRow;
            int SumQty = 0;
            string sQry;
            string sSeq;
            double SumWeight = 0;
            double DocTotal = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sRow = oRow;
                switch (oUID)
                {
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        oMat01.Clear();
                        oMat02.Clear();
                        oDS_PS_MM050L.Clear();
                        oDS_PS_TEMPTABLE.Clear();
                        break;

                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oDS_PS_MM050H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oDS_PS_MM050H.SetValue("U_CntcName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "Mat01":

                        if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                        {
                            oMat01.FlushToDataSource();
                            PS_MM050_AddMatrixRow(oMat01.RowCount, false);
                            oMat01.Columns.Item("PQDocNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        oMat01.FlushToDataSource();

                        sQry = "Select a.DocNum, b.LineId, b.U_ItemCode "; //U_LineNum를 LineId로 수정(2012.07.24 송명규)
                        sQry += "From [@PS_MM010H] a Inner Join [@PS_MM010L] b On a.DocEntry = b.DocEntry ";
                        sQry += "Where a.DocNum = '" + oDS_PS_MM050L.GetValue("U_PQDocNum", oRow - 1).ToString().Trim() + "' ";
                        sQry += "And a.Status = 'O'";
                        oRecordSet.DoQuery(sQry);

                        while (!oRecordSet.EoF)
                        {
                            sSeq = "Y";
                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (oDS_PS_MM050L.GetValue("U_PQDocNum", i).ToString().Trim() == oRecordSet.Fields.Item(0).Value.ToString().Trim() && oDS_PS_MM050L.GetValue("U_PQLinNum", i).ToString().Trim() == oRecordSet.Fields.Item(1).Value.ToString().Trim())
                                {
                                    sSeq = "N";
                                }
                            }
                            if (sSeq == "Y")
                            {
                                oDS_PS_MM050L.SetValue("U_PQDocNum", sRow - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                                oDS_PS_MM050L.SetValue("U_PQLinNum", sRow - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                PS_MM050_AddMatrixRow(sRow, false);
                                sRow += 1;
                            }
                            oRecordSet.MoveNext();
                        }

                        if (oMat01.VisualRowCount > 0)
                        {
                            if (string.IsNullOrEmpty(oDS_PS_MM050L.GetValue("U_ItemCode", oMat01.VisualRowCount - 1).ToString().Trim()))
                            {
                                oDS_PS_MM050L.RemoveRecord(oMat01.VisualRowCount - 1);
                            }
                        }
                        oMat01.LoadFromDataSource();
                        break;

                    case "Mat02":
                        if (oCol == "Weight")
                        {
                            oMat02.FlushToDataSource();
                            if (Convert.ToDouble(oMat02.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()) == 0)
                            {
                                oDS_PS_MM050L.SetValue("U_LinTotal", oRow - 1, "0");
                            }
                            else
                            {
                                //금액 반올림 2012/04/03 노근용 수정
                                oDS_PS_MM050L.SetValue("U_LinTotal", oRow - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat02.Columns.Item("Price").Cells.Item(oRow).Specific.Value.ToString().Trim()), 0)));
                            }
                            oDS_PS_MM050L.SetValue("U_RealWt", oRow - 1, oMat02.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim());
                            oMat02.LoadFromDataSource();
                            oMat02.Columns.Item("Price").Cells.Item(oRow).Click();

                            for (i = 0; i < oMat02.VisualRowCount; i++)
                            {
                                if (string.IsNullOrEmpty(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                {
                                }
                                else
                                {
                                    SumQty += Convert.ToDouble(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                }
                                SumWeight += Convert.ToDouble(oMat02.Columns.Item("RealWt").Cells.Item(i + 1).Specific.Value);
                                DocTotal += Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(i + 1).Specific.Value);
                            }
                            oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                        }
                        else if (oCol == "LinTotal")
                        {
                            oMat02.FlushToDataSource();
                            if (Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(oRow).Specific.Value.ToString().Trim()) == 0)
                            {
                                oDS_PS_MM050L.SetValue("U_Price", oRow - 1, "0");
                            }
                            else
                            {
                                if (Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()) == 0)
                                {
                                    oDS_PS_MM050L.SetValue("U_Price", oRow - 1, "0");
                                }
                                else
                                {
                                    oDS_PS_MM050L.SetValue("U_Price", oRow - 1, Convert.ToString(Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(oRow).Specific.Value.ToString().Trim()) / Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim())));
                                }
                            }
                            oMat02.LoadFromDataSource();

                            for (i = 0; i < oMat02.VisualRowCount; i++)
                            {
                                DocTotal += Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(i + 1).Specific.Value);
                                if (string.IsNullOrEmpty(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                {
                                }
                                else
                                {
                                    SumQty += Convert.ToDouble(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                }
                                SumWeight += Convert.ToDouble(oMat02.Columns.Item("RealWt").Cells.Item(i + 1).Specific.Value);
                            }
                            oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                            oMat02.Columns.Item("Price").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "RealWt")
                        {
                            for (i = 0; i < oMat02.VisualRowCount; i++)
                            {
                                DocTotal += Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(i + 1).Specific.Value);
                                if (string.IsNullOrEmpty(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                {
                                }
                                else
                                {
                                    SumQty += Convert.ToDouble(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                }
                                SumWeight += Convert.ToDouble(oMat02.Columns.Item("RealWt").Cells.Item(i + 1).Specific.Value);
                            }
                            oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                        }
                        else if (oCol == "WhsCode")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oMat02.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);

                            oMat02.Columns.Item("WhsName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }
                        break;
                }
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM050_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM050_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        PS_MM050_Delete_EmptyRow(); //검수입고 문서만 등록 시 이 행은 주석 제외
                        oFormMode01 = oForm.Mode;
                    }
                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        oFormMode01 = oForm.Mode;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                            else if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                PS_MM050_FormItemEnabled();
                                oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true) // 황영수 2018.07.20 Call Sbo_Application.ActivateMenuItem("1282") 오류로 Menu Event 내에 있는 것을 가져옴.
                        {
                            PS_MM050_Initialization();
                            PS_MM050_FormItemEnabled();
                            PS_MM050_FormClear();
                            oDS_PS_MM050H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (PS_MM050_HeaderSpaceLineDel() == false)
                        {
                            oMat01.Clear();
                            oMat02.Clear();
                            oDS_PS_TEMPTABLE.Clear();
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            PS_MM050_Display_MatrixData();
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
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("PQDocNum").Cells.Item(pVal.Row).Specific.Value))
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "BPLId" || pVal.ItemUID == "POType")
                        {
                            oMat01.Clear();
                            oMat02.Clear();
                            oDS_PS_MM050L.Clear();
                            oDS_PS_TEMPTABLE.Clear();
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
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "PODocNum")
                    {
                        PS_MM030 pS_MM030 = new PS_MM030();
                        pS_MM030.LoadForm(oMat01.Columns.Item("PODocNum").Cells.Item(pVal.Row).Specific.Value);
                        pS_MM030 = null;
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            string ItemCode;
            int Qty;
            double Calculate_Weight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            PS_MM050_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM050_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "Qty")
                            {
                                oMat02.FlushToDataSource();
                                ItemCode = oDS_PS_MM050L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                Qty = Convert.ToInt32(oDS_PS_MM050L.GetValue("U_Qty", pVal.Row - 1));

                                Calculate_Weight = dataHelpClass.Calculate_Weight(ItemCode, Qty, oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());

                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    if (oDS_PS_TEMPTABLE.GetValue("U_sField12", pVal.Row - 1).ToString().Trim() == "102" && (oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "30" || oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "40"))
                                    {
                                        //기계사업부 외주가공, 외주제작 중량은 수량과 동일하다.
                                        Calculate_Weight = Qty;
                                        oMat02.Columns.Item("RealWt").Cells.Item(pVal.Row).Specific.Value = Calculate_Weight;
                                    }
                                }
                                else
                                {
                                    if (oDS_PS_MM050L.GetValue("U_WhsCode", pVal.Row - 1) == "102" && (oDS_PS_MM050L.GetValue("U_ItemGpCd", pVal.Row - 1) == "30" || oDS_PS_MM050L.GetValue("U_ItemGpCd", pVal.Row - 1) == "40"))
                                    {
                                        Calculate_Weight = Qty;
                                        oMat02.Columns.Item("RealWt").Cells.Item(pVal.Row).Specific.Value = Calculate_Weight;
                                    }
                                }
                                oMat02.Columns.Item("Weight").Cells.Item(pVal.Row).Specific.Value = Calculate_Weight;
                                oMat02.Columns.Item("RealWt").Cells.Item(pVal.Row).Specific.Value = Calculate_Weight;

                                sQry = "Select ItmsGrpCod From OITM Where ItemCode = '" + ItemCode + "'";
                                oRecordSet.DoQuery(sQry);

                                //부자재
                                if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "105")
                                {
                                    oMat02.Columns.Item("RealWt").Cells.Item(pVal.Row).Specific.Value = Calculate_Weight;
                                }
                            }
                            else if (pVal.ColUID == "Price" || pVal.ColUID == "Weight" || pVal.ColUID == "RealWt")
                            {
                                PS_MM050_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "LinTotal")
                            {
                                PS_MM050_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                PS_MM050_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            int SumQty = 0;
            double DocTotal = 0;
            double SumWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (i = 0; i < oMat02.VisualRowCount; i++)
                    {
                        DocTotal += Convert.ToDouble(oMat02.Columns.Item("LinTotal").Cells.Item(i + 1).Specific.Value);
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                        {
                        }
                        else
                        {
                            SumQty += Convert.ToDouble(oMat02.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                        }
                        SumWeight += Convert.ToDouble(oMat02.Columns.Item("RealWt").Cells.Item(i + 1).Specific.Value);
                    }
                    oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
                    oForm.Items.Item("SumQty").Specific.Value = SumQty;
                    oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                    oMat01.AutoResizeColumns();
                    oMat02.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM050H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM050L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_TEMPTABLE);
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
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            int qField01;
            int iField04;
            int loopCount;
            string selectedValue;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row != 0)
                    {
                        if (oMat02.VisualRowCount == 0)
                        {
                            oDS_PS_MM050L.Clear();
                        }

                        //품의금액 VS 가입고 금액 비교(상단 메트릭스에 값을 조회 해놓은 상태에서 품의금액을 변경한 경우 체크)
                        for (i = 0; i < oMat02.VisualRowCount; i++)
                        {
                            if (oDS_PS_MM050L.GetValue("U_PODocNum", i).ToString().Trim() == oDS_PS_TEMPTABLE.GetValue("U_sField01", pVal.Row - 1).ToString().Trim() & oDS_PS_MM050L.GetValue("U_POLinNum", i).ToString().Trim() == oDS_PS_TEMPTABLE.GetValue("U_iField02", pVal.Row - 1).ToString().Trim())
                            {
                                errMessage = "같은 행을 두번 선택할 수 없습니다. 확인하세요.";
                                throw new Exception();
                                //j = 1;
                            }
                        }

                        oMat02.FlushToDataSource();
                        if (oMat02.VisualRowCount == 0)
                        {
                            PS_MM050_AddMatrixRow02(0, true);
                        }
                        else
                        {
                            PS_MM050_AddMatrixRow02(oMat02.VisualRowCount, false);
                        }
                        oMat02.Columns.Item("PODocNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField01", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("POLinNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_iField02", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("ItemCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField03", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("ItemName").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField04", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("OutSize").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField23", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("OutUnit").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField24", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("ItemGpCd").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim());
                        oMat02.Columns.Item("POType").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_TEMPTABLE.GetValue("U_sField05", pVal.Row - 1).ToString().Trim());
                        qField01 = Convert.ToInt32(Convert.ToDouble(oDS_PS_TEMPTABLE.GetValue("U_qField01", pVal.Row - 1).ToString().Trim()));
                        iField04 = Convert.ToInt32(Convert.ToDouble(oDS_PS_TEMPTABLE.GetValue("U_iField04", pVal.Row - 1).ToString().Trim()));

                        //입고수량과 미입고 수량이 같을때 미입고 수량을 입고수량에 넣어준다
                        if (qField01 == iField04)
                        {
                            if (oDS_PS_TEMPTABLE.GetValue("U_sField12", pVal.Row - 1).ToString().Trim() == "102" && (oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "30" || oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "40"))
                            {
                                oMat02.Columns.Item("Qty").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField01", pVal.Row - 1).ToString().Trim();
                            }
                            else
                            {
                                oMat02.Columns.Item("Qty").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_iField03", pVal.Row - 1).ToString().Trim();
                            }
                            oMat02.Columns.Item("Weight").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField01", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("RealWt").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField01", pVal.Row - 1).ToString().Trim();
                        }
                        oMat02.Columns.Item("LinTotal").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_aField02", pVal.Row - 1).ToString().Trim();
                        oMat02.FlushToDataSource();
                        oDS_PS_MM050L.SetValue("U_Price", oMat02.VisualRowCount - 1, oDS_PS_TEMPTABLE.GetValue("U_aField01", pVal.Row - 1).ToString().Trim());
                        oMat02.LoadFromDataSource();

                        oMat02.Columns.Item("UnWeight").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField02", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("DocCur").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField17", pVal.Row - 1).ToString().Trim(); //통화
                        oMat02.Columns.Item("DocRate").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField04", pVal.Row - 1).ToString().Trim(); //환율
                        oMat02.Columns.Item("FCPrice").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_qField03", pVal.Row - 1).ToString().Trim(); //외화환산단가
                        oMat02.Columns.Item("FCAmount").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_aField03", pVal.Row - 1).ToString().Trim(); //외화환산금액
                        oMat02.Columns.Item("WhsCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField12", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("WhsName").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField13", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("Auto").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_TEMPTABLE.GetValue("U_sField25", pVal.Row - 1).ToString().Trim());

                        //가공비품의(30), 외주제작품의(40) (2011.10.28 송명규 수정)
                        if (oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "30" || oDS_PS_TEMPTABLE.GetValue("U_sField18", pVal.Row - 1).ToString().Trim() == "40")
                        {
                            oMat02.Columns.Item("QEYesNo").Cells.Item(oMat02.VisualRowCount).Specific.Select("Y"); //가공비품의(30), 외주제작품의(40) 외
                        }
                        else
                        {
                            oMat02.Columns.Item("QEYesNo").Cells.Item(oMat02.VisualRowCount).Specific.Select("N");
                        }
                        oMat02.Columns.Item("BDocNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField15", pVal.Row - 1).ToString().Trim();
                        oMat02.Columns.Item("BLineNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_sField16", pVal.Row - 1).ToString().Trim();



                        if (qField01 != iField04)
                        {
                            oMat02.Columns.Item("Weight").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_TEMPTABLE.GetValue("U_iField04", pVal.Row - 1).ToString().Trim();
                        }
                        oMat02.Columns.Item("Qty").Cells.Item(oMat02.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oMat02.AutoResizeColumns();
                        BubbleEvent = false;
                    }
                    else if (pVal.ItemUID == "Mat01" && pVal.Row == 0) //더블클릭 열 정렬
                    {
                        oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "QEYesNo")
                        {
                            if (pVal.Row == 0 && oMat02.RowCount > 1)
                            {
                                selectedValue = oMat02.Columns.Item("QEYesNo").Cells.Item(1).Specific.Selected.Value; //첫 행에 선택된 값을 저장

                                for (loopCount = 1; loopCount < oMat02.VisualRowCount; loopCount++)
                                {
                                    oMat02.Columns.Item("QEYesNo").Cells.Item(loopCount).Specific.Select(selectedValue);
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row == 0) //더블클릭 열 정렬
                    {
                        oMat01.FlushToDataSource();
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oDS_PS_TEMPTABLE.SetValue("U_iField01", i - 1, Convert.ToString(i));
                        }
                        oMat01.LoadFromDataSource();
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
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// RESIZE 이벤트
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
                    PS_MM050_FormResize();
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

                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_MM050_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat02.VisualRowCount; i++)
                        {
                            oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat02.FlushToDataSource();
                        oDS_PS_MM050L.RemoveRecord(oDS_PS_MM050L.Size - 1);
                        oMat02.LoadFromDataSource();
                        oForm.Update();
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
                                if (PS_MM050_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1"))
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
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_MM050_Validate("닫기") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 닫기하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1"))
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 닫기할수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            PS_MM050_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_MM050_FormItemEnabled();
                            oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
                            if (dataHelpClass.User_SuperUserYN() == "N") //수퍼유저인 경우는 사번 미표기
                            {
                                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                            }
                            oForm.Items.Item("POStatus").Specific.Select("Y");
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;

                        case "1282": //추가
                            PS_MM050_Initialization();
                            PS_MM050_FormItemEnabled();
                            PS_MM050_FormClear();
                            oDS_PS_MM050H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                            break;

                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            oMat01.FlushToDataSource();
                            oMat01.Clear();
                            oDS_PS_TEMPTABLE.Clear();
                            oMat01.LoadFromDataSource();
                            PS_MM050_FormItemEnabled();
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

                if (pVal.ItemUID == "Mat02")
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
    }
}
