using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 자산이력관리
    /// </summary>
    internal class PS_FX010 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_FX010H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_FX010L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oBPLId;
        private string oDocdate;
        private string oHisType;
        private string oClasCode;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX010.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_FX010_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_FX010");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_FX010_CreateItems();
                PS_FX010_ComboBox_Setting();
                PS_FX010_AddMatrixRow(0, true);
                PS_FX010_LoadCaption();
                PS_FX010_FormClear();

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
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
        private void PS_FX010_CreateItems()
        {
            try
            {
                oDS_PS_FX010H = oForm.DataSources.DBDataSources.Item("@PS_FX010H");
                oDS_PS_FX010L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장_S
                oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");
                //사업장_E

                //이력구분_S
                oForm.DataSources.UserDataSources.Add("SHisType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SHisType").Specific.DataBind.SetBound(true, "", "SHisType");
                //이력구분_E

                //자산코드(검색용)_S
                oForm.DataSources.UserDataSources.Add("CFixCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CFixCode").Specific.DataBind.SetBound(true, "", "CFixCode");
                //자산코드(검색용)_E

                //이력일자_From_S
                oForm.DataSources.UserDataSources.Add("DocDateF", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDateF").Specific.DataBind.SetBound(true, "", "DocDateF");
                //이력일자_From_E

                //이력일자_To_S
                oForm.DataSources.UserDataSources.Add("DocDateT", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("DocDateT").Specific.DataBind.SetBound(true, "", "DocDateT");
                //이력일자_To_E

                oForm.Items.Item("DocDateF").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
                oForm.Items.Item("DocDateT").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_FX010_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //이력구분
                oForm.Items.Item("HisType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("HisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", false, false);
                oForm.Items.Item("HisType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //자산분류
                oForm.Items.Item("ClasCode").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ClasCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", false, false);
                oForm.Items.Item("ClasCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //이력구분(조회조건)
                oForm.Items.Item("SHisType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SHisType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", false, false);
                oForm.Items.Item("SHisType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("AmtYN").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("AmtYN").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("AmtYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", ""); //사업장(매트릭스)
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("HisType"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX002'", "", ""); //이력구분(매트릭스)
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ClasCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", ""); //자산구분(매트릭스)
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PS_FX010_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
                    oForm.Items.Item("BtnDelete").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                    oForm.Items.Item("BtnDelete").Enabled = true;
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
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_FX010_HeaderSpaceLineDel()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (oForm.Items.Item("HisType").Specific.Value.ToString().Trim() == "%")
                {
                    errMessage = "";
                    throw new Exception();
                }
                else if (oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim() == "%")
                {
                    errMessage = "자산구분은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FixCode").Specific.Value.ToString().Trim()))
                {
                    errMessage = "자산코드는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FixName").Specific.Value.ToString().Trim()))
                {
                    errMessage = "자산명은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
                {
                    errMessage = "이력일자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("AmtYN").Specific.Value.ToString().Trim() == "Y")
                {
                    if (Convert.ToDouble(oForm.Items.Item("Amt").Specific.Value) == 0)
                    {
                        errMessage = "자본적지출시 금액은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
                {
                    if (oForm.Items.Item("HisType").Specific.Value.ToString().Trim() == "03")
                    {
                        errMessage = "매각등록시 거래처는 필수사항입니다. 확인하세요.";
                        throw new Exception();
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
            return ReturnValue;
        }

        /// <summary>
        /// PS_FX010_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_FX010_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_FX010L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_FX010L.Offset = oRow;
                oDS_PS_FX010L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// PS_FX010_MTX01
        /// </summary>
        private void PS_FX010_MTX01()
        {
            string errMessage = string.Empty;
            int i;
            string sQry;
            string SBPLID; //사업장
            string SHisType; //이력구분
            string DocDateF;
            string DocDateT;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
                SHisType = oForm.Items.Item("SHisType").Specific.Value.ToString().Trim();
                DocDateF = oForm.Items.Item("DocDateF").Specific.Value.ToString().Trim();
                DocDateT = oForm.Items.Item("DocDateT").Specific.Value.ToString().Trim();

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                sQry = "EXEC [PS_FX010_01] '" + SBPLID + "','" + DocDateF + "','" + DocDateT + "','" + SHisType + "'";
                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_FX010L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PS_FX010_AddMatrixRow(0, true);
                    PS_FX010_LoadCaption();
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_FX010L.Size)
                    {
                        oDS_PS_FX010L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_FX010L.Offset = i;
                    oDS_PS_FX010L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_FX010L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value).ToString("yyyyMMdd"));
                    oDS_PS_FX010L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("HIsType").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ClasCode").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("AmtYN").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("FixCode").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("SubCode").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("FixName").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("Qty").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amt").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Comments").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_FX010L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());

                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_FX010_FormClear()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                {
                    oDS_PS_FX010H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_FX010H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToInt32(oRecordSet.Fields.Item(0).Value) + 1));
                }
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
        /// PS_FX010_DeleteData
        /// </summary>
        private void PS_FX010_DeleteData()
        {
            string errMessage = string.Empty;
            string sQry ;
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                    sQry = "SELECT COUNT(*) FROM [@PS_FX010H] WHERE DocEntry = '" + DocEntry + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount == 0)
                    {
                        errMessage = "삭제대상이 없습니다. 확인하세요.";
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        throw new Exception();
                    }
                    else
                    {
                        sQry = "DELETE FROM [@PS_FX010H] WHERE DocEntry = '" + DocEntry + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                }
                dataHelpClass.MDC_GF_Message("삭제 완료!", "S");
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
        }

        /// <summary>
        /// PS_FX010_UpdateData
        /// </summary>
        private bool PS_FX010_UpdateData()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;
            int DocEntry;
            string sQry;
            string BPLId; //사업장
            string HisType; //이력구분
            string ClasCode; //자산구분
            string DocDate; //이력일자
            string FixCode; //자산코드
            string SubCode; //자산순번
            string FixName; //자산명
            string AmtYN; //자본적지출여부
            string Qty; //수량
            string Amt; //금액
            string Comments; //비고사항
            string CardCode; //거래처
            string CardName; //거래처명
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = Convert.ToInt16(oForm.Items.Item("DocEntry").Specific.Value);
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                HisType = oForm.Items.Item("HisType").Specific.Value.ToString().Trim(); //이력구분
                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim(); //자산구분
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim(); //이력일자
                AmtYN = oForm.Items.Item("AmtYN").Specific.Value.ToString().Trim(); //자본적지출여부
                FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();  //자산코드
                SubCode = oForm.Items.Item("SubCode").Specific.Value.ToString().Trim(); //자산순번
                FixName = oForm.Items.Item("FixName").Specific.Value.ToString().Trim(); //자산명
                Qty = oForm.Items.Item("Qty").Specific.Value.ToString().Trim(); //수량
                Amt = oForm.Items.Item("Amt").Specific.Value.ToString().Trim(); //금액
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim(); //비고
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim(); //거래처
                CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim(); //거래처명

                if (string.IsNullOrEmpty(Convert.ToString(DocEntry)))
                {
                    errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!";
                    throw new Exception();
                }

                sQry = "        UPDATE   [@PS_FX010H]";
                sQry += " SET      U_BPLId = '" + BPLId + "',";
                sQry += "          U_HisType = '" + HisType + "',";
                sQry += "          U_ClasCode = '" + ClasCode + "',";
                sQry += "          U_DocDate = '" + DocDate + "',";
                sQry += "          U_AmtYN = '" + AmtYN + "',";
                sQry += "          U_FixCode = '" + FixCode + "',";
                sQry += "          U_SubCode = '" + SubCode + "',";
                sQry += "          U_FixName = '" + FixName + "',";
                sQry += "          U_Qty  = '" + Qty + "',";
                sQry += "          U_Amt  = '" + Amt + "',";
                sQry += "          U_Comments = '" + Comments + "',";
                sQry += "          U_CardCode = '" + CardCode + "',";
                sQry += "          U_CardName = '" + CardName + "'";
                sQry += " WHERE    DocEntry = '" + DocEntry + "'";

                oRecordSet.DoQuery(sQry);

                dataHelpClass.MDC_GF_Message("수정 완료!", "S");
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
        /// PS_FX010_UpdateData
        /// </summary>
        private bool PS_FX010_AddData()
        {
            bool ReturnValue = false;
            int DocEntry;
            string sQry;
            string BPLId; //사업장
            string HisType; //이력구분
            string ClasCode; //자산구분
            string DocDate;  //이력일자
            string FixCode; //자산코드
            string SubCode; //자산순번
            string FixName; //자산명
            string AmtYN; //자본적지출여부
            string Qty; //수량
            string Amt; //금액
            string Comments; //비고사항
            string CardCode; //거래처
            string CardName; //거래처명
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); //사업장
                HisType = oForm.Items.Item("HisType").Specific.Value.ToString().Trim(); //이력구분
                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim(); //자산구분
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim(); //이력일자
                AmtYN = oForm.Items.Item("AmtYN").Specific.Value.ToString().Trim(); //자본적지출여부
                FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim(); //자산코드
                SubCode = oForm.Items.Item("SubCode").Specific.Value.ToString().Trim(); //자산순번
                FixName = oForm.Items.Item("FixName").Specific.Value.ToString().Trim(); //자산명
                Qty = oForm.Items.Item("Qty").Specific.Value.ToString().Trim(); //수량
                Amt = oForm.Items.Item("Amt").Specific.Value.ToString().Trim(); //금액
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim(); //비고
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim(); //거래처
                CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim(); //거래처명

                //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_FX010H]";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    DocEntry = 1;
                }
                else
                {
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item(0).Value) + 1;
                }

                sQry = " INSERT INTO [@PS_FX010H]";
                sQry += " (";
                sQry += "     DocEntry,";
                sQry += "     DocNum,";
                sQry += "     U_BPLId,";
                sQry += "     U_HIsType,";
                sQry += "     U_ClasCode,";
                sQry += "     U_DocDate,";
                sQry += "     U_FixCode,";
                sQry += "     U_SubCode,";
                sQry += "     U_FixName,";
                sQry += "     U_AmtYN,";
                sQry += "     U_Qty,";
                sQry += "     U_Amt,";
                sQry += "     U_Comments,";
                sQry += "     U_CardCode,";
                sQry += "     U_CardName";
                sQry += " )";
                sQry += " VALUES";
                sQry += " (";
                sQry += DocEntry + ",";
                sQry += DocEntry + ",";
                sQry += "'" + BPLId + "',";
                sQry += "'" + HisType + "',";
                sQry += "'" + ClasCode + "',";
                sQry += "'" + DocDate + "',";
                sQry += "'" + FixCode + "',";
                sQry += "'" + SubCode + "',";
                sQry += "'" + FixName + "',";
                sQry += "'" + AmtYN + "',";
                sQry += "'" + Qty + "',";
                sQry += "'" + Amt + "',";
                sQry += "'" + Comments + "',";
                sQry += "'" + CardCode + "',";
                sQry += "'" + CardName + "'";
                sQry += ")";

                oRecordSet02.DoQuery(sQry);

                dataHelpClass.MDC_GF_Message("등록 완료!", "S");
                ReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            }
            return ReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_FX010_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            string FixCode;
            string SubCode;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "CFixCode":
                        FixCode = codeHelpClass.Left(oForm.Items.Item("CFixCode").Specific.Value, 6);
                        SubCode = codeHelpClass.Right(oForm.Items.Item("CFixCode").Specific.Value, 3);
                        oForm.Items.Item("FixCode").Specific.Value = FixCode;
                        oForm.Items.Item("SubCode").Specific.Value = SubCode;

                        sQry = "Select U_FixName From [@PS_FX005H] Where U_FixCode = '" + FixCode + "'";
                        sQry += " and U_SubCode = '" + SubCode + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("FixName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "SubCode": //자산코드
                        break;

                    case "CardCode": //거래처
                        oDS_PS_FX010H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("CardCode").Specific.Value + "'", ""));
                        break;
                }
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
        /// PS_FX010_FormReset
        /// </summary>
        private void PS_FX010_FormReset()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FX010H]";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
                {
                    oDS_PS_FX010H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_FX010H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
                }

                if (string.IsNullOrEmpty(oBPLId))
                {
                    oBPLId = dataHelpClass.User_BPLID();
                }
                if (string.IsNullOrEmpty(oHisType))
                {
                    oHisType = "%";
                }
                if (string.IsNullOrEmpty(oClasCode))
                {
                    oClasCode = "%";
                }
                if (string.IsNullOrEmpty(oDocdate))
                {
                    oDocdate = oDocdate;
                }

                oDS_PS_FX010H.SetValue("U_BPLId", 0, oBPLId); //사업장
                oDS_PS_FX010H.SetValue("U_HisType", 0, oHisType); //이력구분
                oDS_PS_FX010H.SetValue("U_ClasCode", 0, oClasCode); //자산분류
                oDS_PS_FX010H.SetValue("U_AmtYN", 0, "N"); //자본적지출여부
                oDS_PS_FX010H.SetValue("U_DocDate", 0, oDocdate); //이력일자
                oDS_PS_FX010H.SetValue("U_FixCode", 0, ""); //자산코드
                oDS_PS_FX010H.SetValue("U_SubCode", 0, ""); //자산순번
                oDS_PS_FX010H.SetValue("U_FixName", 0, ""); //자산명
                oDS_PS_FX010H.SetValue("U_Qty", 0, "0"); //수량
                oDS_PS_FX010H.SetValue("U_Amt", 0, "0"); //금액
                oDS_PS_FX010H.SetValue("U_Comments", 0, ""); //비고사항
                oForm.Items.Item("CFixCode").Specific.Value = "";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnAdd") //추가/확인 버튼클릭
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_FX010_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_FX010_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_FX010_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_FX010_LoadCaption();
                            PS_FX010_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                            if (PS_FX010_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_FX010_UpdateData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_FX010_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_FX010_LoadCaption();
                            PS_FX010_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch")
                    {

                        PS_FX010_FormReset();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        PS_FX010_LoadCaption();
                        PS_FX010_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete")
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == 1)
                        {

                            PS_FX010_DeleteData();
                            PS_FX010_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_FX010_LoadCaption();
                            PS_FX010_MTX01();
                        }
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CFixCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CFixCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "BPLId")
                    {
                        oBPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                    }
                    else if (pVal.ItemUID == "ClasCode")
                    {
                        oClasCode = oForm.Items.Item("ClasCode").Specific.Value;
                    }
                    else if (pVal.ItemUID == "ClasCode")
                    {
                        oHisType = oForm.Items.Item("HisType").Specific.Value;
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oDS_PS_FX010H.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value); //관리번호
                            oDS_PS_FX010H.SetValue("U_BPLId", 0, oMat01.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Value); //사업장
                            oDS_PS_FX010H.SetValue("U_HisType", 0, oMat01.Columns.Item("HisType").Cells.Item(pVal.Row).Specific.Value); //이력구분
                            oDS_PS_FX010H.SetValue("U_ClasCode", 0, oMat01.Columns.Item("ClasCode").Cells.Item(pVal.Row).Specific.Value); //자산분류
                            oDS_PS_FX010H.SetValue("U_AmtYN", 0, oMat01.Columns.Item("AmtYN").Cells.Item(pVal.Row).Specific.Value); //자본적지출여부
                            oDS_PS_FX010H.SetValue("U_DocDate", 0, oMat01.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value); //이력일자
                            oDS_PS_FX010H.SetValue("U_FixCode", 0, oMat01.Columns.Item("FixCode").Cells.Item(pVal.Row).Specific.Value); //자산코드
                            oDS_PS_FX010H.SetValue("U_SubCode", 0, oMat01.Columns.Item("SubCode").Cells.Item(pVal.Row).Specific.Value); //자산순번
                            oDS_PS_FX010H.SetValue("U_FixName", 0, oMat01.Columns.Item("FixName").Cells.Item(pVal.Row).Specific.Value); //자산명
                            oDS_PS_FX010H.SetValue("U_Qty", 0, oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value); //수량
                            oDS_PS_FX010H.SetValue("U_Amt", 0, oMat01.Columns.Item("Amt").Cells.Item(pVal.Row).Specific.Value); //금액
                            oDS_PS_FX010H.SetValue("U_Comments", 0, oMat01.Columns.Item("Comments").Cells.Item(pVal.Row).Specific.Value); //비고
                            oDS_PS_FX010H.SetValue("U_CardCode", 0, oMat01.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.Value); //거래처
                            oDS_PS_FX010H.SetValue("U_CardName", 0, oMat01.Columns.Item("CardName").Cells.Item(pVal.Row).Specific.Value); //거래처명
                            oForm.Items.Item("CFixCode").Specific.Value = oMat01.Columns.Item("FixCode").Cells.Item(pVal.Row).Specific.Value + "-" + oMat01.Columns.Item("SubCode").Cells.Item(pVal.Row).Specific.Value;

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_FX010_LoadCaption();
                        }
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            PS_FX010_FlushToItemValue(pVal.ItemUID, 0, "");
                            if (pVal.ItemUID == "BPLId")
                            {
                                oBPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                            }
                            else if (pVal.ItemUID == "DocDate")
                            {
                                oDocdate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                            }
                            else if (pVal.ItemUID == "HisType")
                            {
                                oHisType = oForm.Items.Item("HisType").Specific.Value;
                            }
                            else if (pVal.ItemUID == "ClasCode")
                            {
                                oClasCode = oForm.Items.Item("ClasCode").Specific.Value;
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX010H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX010L);
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
            try
            {
                int i = 0;
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true) 
                    { 
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_FX010H.RemoveRecord(oDS_PS_FX010H.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_FX010_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_FX010H.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_FX010_AddMatrixRow(oMat01.RowCount, false);
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
                            PS_FX010_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PS_FX010_LoadCaption();
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
