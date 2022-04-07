using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.IO;
using Scripting;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 고정자산등록
    /// </summary>
    internal class PS_FX005 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_FX005H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_FX005L; //등록라인

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX005.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_FX005_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_FX005");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_FX005_CreateItems();
                PS_FX005_ComboBox_Setting();
                PS_FX005_AddMatrixRow(0, true);
                PS_FX005_LoadCaption();
                PS_FX005_Initialization();

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
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
        private void PS_FX005_CreateItems()
        {
            try
            {
                oDS_PS_FX005H = oForm.DataSources.DBDataSources.Item("@PS_FX005H");
                oDS_PS_FX005L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oForm.Items.Item("FixCode").Enabled = false;

                oMat01 = oForm.Items.Item("Mat01").Specific;
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
        private void PS_FX005_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Items.Item("SubDiv").Specific.ValidValues.Add("N", "메인자산");
                oForm.Items.Item("SubDiv").Specific.ValidValues.Add("Y", "Sub자산");
                oForm.Items.Item("SubDiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                dataHelpClass.Set_ComboList(oForm.Items.Item("ClasCode").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor", "", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("SClasCod").Specific, "Select U_Minor = '%', U_CdName = '전체' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor", "", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("DepCode").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX003' order by U_Minor", "", false, false);
                oForm.Items.Item("SClasCod").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                sQry = "SELECT BPLId, BPLName FROM OBPL order by BPLId";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX001' order by U_Minor";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oMat01.Columns.Item("ClasCode").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
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
        /// Initialization
        /// </summary>
        private void PS_FX005_Initialization()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = Convert.ToInt32(oRecordSet.Fields.Item(0).Value) + 1;
                }
                oForm.Items.Item("PostDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
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
        private bool PS_FX005_HeaderSpaceLineDel()
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim()))
                {
                    errMessage = "자산분류는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FixCode").Specific.Value.ToString().Trim()))
                {
                    errMessage = "자산코드는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("FixName").Specific.Value.ToString().Trim()))
                {
                    errMessage = "자산명은 필수사항입니다. 확인하세요";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("PostDate").Specific.Value.ToString().Trim()))
                {
                    errMessage = "취득일자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (Convert.ToDouble(oForm.Items.Item("PostQty").Specific.Value.ToString().Trim()) == 0)
                {
                    errMessage = "취득수량은 필수사항입니다. 확인하세요.";
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
        /// PS_FX005_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_FX005_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_FX005L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_FX005L.Offset = oRow;
                oDS_PS_FX005L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

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
        /// PS_FX005_MTX01
        /// </summary>
        private void PS_FX005_MTX01(string pFixCode)
        {
            short i;
            string sQry;
            string SBPLID;
            string SClasCod;
            string SFixCode;
            string SFixName;
            string STempChr1;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                SFixCode = pFixCode;
                SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
                SClasCod = oForm.Items.Item("SClasCod").Specific.Value.ToString().Trim();
                SFixName = oForm.Items.Item("SFixName").Specific.Value.ToString().Trim();
                STempChr1 = oForm.Items.Item("STempChr1").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(SClasCod))
                {
                    SClasCod = "%";
                }
                if (string.IsNullOrEmpty(SFixCode))
                {
                    SFixCode = "%";
                }
                if (string.IsNullOrEmpty(SFixName))
                {
                    SFixName = "%";
                }
                if (string.IsNullOrEmpty(STempChr1))
                {
                    STempChr1 = "%";
                }

                sQry = " EXEC [PS_FX005_01] '";
                sQry += SBPLID + "','";
                sQry += SClasCod + "', '";
                sQry += SFixCode + "', '";
                sQry += SFixName + "', '";
                sQry += STempChr1 + "'";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_FX005L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W");
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PS_FX005_AddMatrixRow(0, true);
                    PS_FX005_LoadCaption();
                    throw new Exception();
                }

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_FX005L.Size)
                    {
                        oDS_PS_FX005L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_FX005L.Offset = i;
                    oDS_PS_FX005L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_FX005L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim()); //BPLId
                    oDS_PS_FX005L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(1).Value.ToString().Trim()); //DocEntry
                    oDS_PS_FX005L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(2).Value.ToString().Trim()); //FixCode
                    oDS_PS_FX005L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(3).Value.ToString().Trim()); //SubCode
                    oDS_PS_FX005L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(4).Value.ToString().Trim()); //FixName
                    oDS_PS_FX005L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(5).Value.ToString().Trim()); //ClasCode
                    oDS_PS_FX005L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(6).Value.ToString().Trim()); //PostDate
                    oDS_PS_FX005L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item(7).Value.ToString().Trim()); //PostQty
                    oDS_PS_FX005L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(8).Value.ToString().Trim()); //PostAmt
                    oDS_PS_FX005L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item(9).Value.ToString().Trim()); //TeamNm
                    oDS_PS_FX005L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item(10).Value.ToString().Trim()); //RspNm
                    oDS_PS_FX005L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item(11).Value.ToString().Trim()); //StopYN
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
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_FX005_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "CardCode": //거래처
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "FixCode":
                        sQry = "Select right('00' + Convert(Nvarchar(3),Convert(Numeric(3,0), Max(U_SubCode)) + 1),3) From [@PS_FX005H] Where U_FixCode = '" + oForm.Items.Item("FixCode").Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("SubCode").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "TeamCode": //팀
                        sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code ";
                        sQry = sQry + " WHERE   T0.Name = '부서' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value + "' ";
                        sQry = sQry + " And T1.U_Code = '" + oForm.Items.Item("TeamCode").Specific.Value + "'";

                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("TeamNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "RspCode": //담당
                        sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code ";
                        sQry = sQry + " WHERE   T0.Name = '담당' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value + "' ";
                        sQry = sQry + " And T1.U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value + "'";
                        sQry = sQry + " And T1.U_Code = '" + oForm.Items.Item("RspCode").Specific.Value + "'";

                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("RspNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        break;

                    case "PrcCode":
                        sQry = "SELECT OCRNAME FROM OOCR WHERE DIMCODE = '1' And OCRCODE = '" + oForm.Items.Item("PrcCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Items.Item("PrcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PS_FX005_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Btn_save").Specific.Caption = "추가";
                    oForm.Items.Item("Btn_del").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("Btn_save").Specific.Caption = "수정";
                    oForm.Items.Item("Btn_del").Enabled = true;
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
        /// PS_FX005_DeleteData
        /// </summary>
        private void PS_FX005_DeleteData()
        {
            string errMessage = string.Empty;
            string sQry;
            string DocEntry;
            string FixCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                    FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();

                    sQry = " Select Count(*) ";
                    sQry += "  From [@PS_FX005H] ";
                    sQry += " where DocEntry = '" + DocEntry + "' ";
                    sQry += "   And U_FixCode = '" + FixCode + "'";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount == 0)
                    {
                        dataHelpClass.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W");
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    }
                    else
                    {
                        sQry = "Delete From [@PS_FX005H] where DocEntry = '" + DocEntry + "' And U_FixCode = '" + FixCode + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                }
                PS_FX005_FormReset();
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.Items.Item("Btn_ret").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// PS_FX005_UpdateData
        /// </summary>
        private bool PS_FX005_UpdateData(SAPbouiCOM.ItemEvent pval)
        {
            bool ReturnValue = false;
            string errMessage = string.Empty;
            string sQry;
            string DocEntry;
            string BPLId;
            string ClasCode;
            string FixCode;
            string FixName;
            string PostDate;
            string PostAmt;
            string PostQty;
            string LongYear;
            string DepCode;
            string DepRate;
            string CardCode;
            string CardName;
            string Kw;
            string Volt;
            string TeamCode;
            string TeamNm;
            string RspCode;
            string RspNm;
            string PrcCode;
            string PrcName;
            string Comments;
            string Place;
            string TempChr1;
            string StopYN;
            string SubDiv;
            string SubCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim();
                SubDiv = oForm.Items.Item("SubDiv").Specific.Value.ToString().Trim();
                FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();
                SubCode = oForm.Items.Item("SubCode").Specific.Value.ToString().Trim();
                FixName = oForm.Items.Item("FixName").Specific.Value.ToString().Trim();
                PostDate = oForm.Items.Item("PostDate").Specific.Value.ToString().Trim();
                PostAmt = oForm.Items.Item("PostAmt").Specific.Value.ToString().Trim();
                PostQty = oForm.Items.Item("PostQty").Specific.Value.ToString().Trim();
                LongYear = oForm.Items.Item("LongYear").Specific.Value.ToString().Trim();
                DepCode = oForm.Items.Item("DepCode").Specific.Value.ToString().Trim();
                DepRate = oForm.Items.Item("DepRate").Specific.Value.ToString().Trim();

                if (oForm.Items.Item("StopYN").Specific.Checked == true)
                {
                    StopYN = "Y";
                }
                else
                {
                    StopYN = "N";
                }

                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
                Kw = oForm.Items.Item("Kw").Specific.Value.ToString().Trim();
                Volt = oForm.Items.Item("Volt").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                TeamNm = oForm.Items.Item("TeamNm").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                RspNm = oForm.Items.Item("RspNm").Specific.Value.ToString().Trim();
                PrcCode = oForm.Items.Item("PrcCode").Specific.Value.ToString().Trim();
                PrcName = oForm.Items.Item("PrcName").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                Place = oForm.Items.Item("Place").Specific.Value.ToString().Trim();
                TempChr1 = oForm.Items.Item("TempChr1").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(DocEntry.ToString().Trim()))
                {
                    errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!";
                    throw new Exception();
                }

                sQry = " Update [@PS_FX005H]";
                sQry += " set ";
                sQry += " U_BPLId = '" + BPLId + "',";
                sQry += " U_ClasCode = '" + ClasCode + "',";
                sQry += " U_SubDiv = '" + SubDiv + "',";
                sQry += " U_FixCode = '" + FixCode + "',";
                sQry += " U_SubCode = '" + SubCode + "',";
                sQry += " U_FixName = '" + FixName + "',";
                sQry += " U_PostDate = '" + PostDate + "',";
                sQry += " U_PostAmt = '" + PostAmt + "',";
                sQry += " U_PostQty  = '" + PostQty + "',";
                sQry += " U_LongYear  = '" + LongYear + "',";
                sQry += " U_DepCode  = '" + DepCode + "',";
                sQry += " U_DepRate = '" + DepRate + "',";
                sQry += " U_StopYN = '" + StopYN + "',";
                sQry += " U_CardCode  = '" + CardCode + "',";
                sQry += " U_CardName  = '" + CardName + "',";
                sQry += " U_Kw  = '" + Kw + "',";
                sQry += " U_Volt  = '" + Volt + "',";
                sQry += " U_TeamCode  = '" + TeamCode + "',";
                sQry += " U_TeamNm  = '" + TeamNm + "',";
                sQry += " U_RspCode  = '" + RspCode + "',";
                sQry += " U_RspNm  = '" + RspNm + "',";
                sQry += " U_PrcCode  = '" + PrcCode + "',";
                sQry += " U_PrcName  = '" + PrcName + "',";
                sQry += " U_Place  = '" + Place + "',";
                sQry += " U_TempChr1  = '" + TempChr1 + "',";
                sQry += " U_Comments  = '" + Comments + "'";
                sQry += " Where DocEntry = '" + DocEntry + "'";
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
        /// PS_FX005_AddData
        /// </summary>
        private bool PS_FX005_AddData(SAPbouiCOM.ItemEvent pVal)
        {
            bool ReturnValue = false;
            string sQry;
            string LongYear;
            string PostAmt;
            string FixName;
            string ClasCode;
            string DocEntry;
            string BPLId;
            string FixCode;
            string PostDate;
            string PostQty;
            string DepCode;
            string PrcName;
            string RspNm;
            string TeamNm;
            string Volt;
            string CardName;
            string DepRate;
            string CardCode;
            string Kw;
            string TeamCode;
            string RspCode;
            string PrcCode;
            string Comments;
            string TempChr1;
            string Place;
            string StopYN;
            string SubDiv;
            string SubCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim();
                SubDiv = oForm.Items.Item("SubDiv").Specific.Value.ToString().Trim();
                FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();
                FixName = oForm.Items.Item("FixName").Specific.Value.ToString().Trim();
                SubCode = oForm.Items.Item("SubCode").Specific.Value.ToString().Trim();
                PostDate = oForm.Items.Item("PostDate").Specific.Value.ToString().Trim();
                PostAmt = oForm.Items.Item("PostAmt").Specific.Value.ToString().Trim();
                PostQty = oForm.Items.Item("PostQty").Specific.Value.ToString().Trim();
                LongYear = oForm.Items.Item("LongYear").Specific.Value.ToString().Trim();
                DepCode = oForm.Items.Item("DepCode").Specific.Value.ToString().Trim();
                DepRate = oForm.Items.Item("DepRate").Specific.Value.ToString().Trim();

                if (oForm.Items.Item("StopYN").Specific.Checked == true)
                {
                    StopYN = "Y";
                }
                else
                {
                    StopYN = "N";
                }
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
                Kw = oForm.Items.Item("Kw").Specific.Value.ToString().Trim();
                Volt = oForm.Items.Item("Volt").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                TeamNm = oForm.Items.Item("TeamNm").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
                RspNm = oForm.Items.Item("RspNm").Specific.Value.ToString().Trim();
                PrcCode = oForm.Items.Item("PrcCode").Specific.Value.ToString().Trim();
                PrcName = oForm.Items.Item("PrcName").Specific.Value.ToString().Trim();
                Comments = oForm.Items.Item("Comments").Specific.Value.ToString().Trim();
                Place = oForm.Items.Item("Place").Specific.Value.ToString().Trim();
                TempChr1 = oForm.Items.Item("TempChr1").Specific.Value.ToString().Trim();

                sQry = "Select right('00' + Convert(Nvarchar(3),Convert(Numeric(3,0), Isnull(Max(U_SubCode),'000')) + 1),3) From [@PS_FX005H] Where U_FixCode = '" + FixCode + "'";
                oRecordSet.DoQuery(sQry);

                SubCode = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]";
                oRecordSet.DoQuery(sQry);
                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                {
                    DocEntry = "1";
                }
                else
                {
                    DocEntry = Convert.ToString(Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
                }

                sQry = "INSERT INTO [@PS_FX005H]";
                sQry += " (";
                sQry += " DocEntry,";
                sQry += " DocNum,";
                sQry += " U_BPLId,";
                sQry += " U_ClasCode,";
                sQry += " U_SubDiv,";
                sQry += " U_FixCode,";
                sQry += " U_SubCode,";
                sQry += " U_FixName,";
                sQry += " U_PostDate,";
                sQry += " U_PostAmt,";
                sQry += " U_PostQty,";
                sQry += " U_LongYear,";
                sQry += " U_DepCode,";
                sQry += " U_DepRate,";
                sQry += " U_StopYN,";
                sQry += " U_CardCode,";
                sQry += " U_CardName,";
                sQry += " U_Kw,";
                sQry += " U_Volt,";
                sQry += " U_TeamCode,";
                sQry += " U_TeamNm,";
                sQry += " U_RspCode,";
                sQry += " U_RspNm,";
                sQry += " U_PrcCode,";
                sQry += " U_PrcName,";
                sQry += " U_Place,";
                sQry += " U_Comments,";
                sQry += " U_TempChr1";
                sQry += " ) ";
                sQry += "VALUES(";
                sQry += DocEntry + ",";
                sQry += DocEntry + ",";
                sQry += "'" + BPLId + "',";
                sQry += "'" + ClasCode + "',";
                sQry += "'" + SubDiv + "',";
                sQry += "'" + FixCode + "',";
                sQry += "'" + SubCode + "',";
                sQry += "'" + FixName + "',";
                sQry += "'" + PostDate + "',";
                sQry += "'" + PostAmt + "',";
                sQry += "'" + PostQty + "',";
                sQry += "'" + LongYear + "',";
                sQry += "'" + DepCode + "',";
                sQry += "'" + DepRate + "',";
                sQry += "'" + StopYN + "',";
                sQry += "'" + CardCode + "',";
                sQry += "'" + CardName + "',";
                sQry += "'" + Kw + "',";
                sQry += "'" + Volt + "',";
                sQry += "'" + TeamCode + "',";
                sQry += "'" + TeamNm + "',";
                sQry += "'" + RspCode + "',";
                sQry += "'" + RspNm + "',";
                sQry += "'" + PrcCode + "',";
                sQry += "'" + PrcName + "',";
                sQry += "'" + Place + "',";
                sQry += "'" + Comments + "',";
                sQry += "'" + TempChr1 + "'";
                sQry += ")";
                oRecordSet02.DoQuery(sQry);

                // 기존 테이블에 값을 추가하면 ZPS_FX005_PIC 테이블 고정자산 사진용으로 추가.
                sQry = "insert into PSHDB_IMG.dbo.ZPS_FX005_PIC(BPLId,FixCode,SubCode) SELECT ";
                sQry += "'" + BPLId + "',";
                sQry += "'" + FixCode + "',";
                sQry += "'" + SubCode + "'";
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
        /// PS_FX005_FormReset
        /// </summary>
        private void PS_FX005_FormReset()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_FX005H]";
                oRecordSet.DoQuery(sQry);
                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = Convert.ToInt32(oRecordSet.Fields.Item(0).Value) + 1;
                }
                oDS_PS_FX005H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
                oDS_PS_FX005H.SetValue("U_ClasCode", 0, "");
                oDS_PS_FX005H.SetValue("U_SubDiv", 0, "N");  //서브자산여부추가(2017.09.05 송명규) (N:메인자산, Y:서브자산)
                oDS_PS_FX005H.SetValue("U_FixCode", 0, "");
                oDS_PS_FX005H.SetValue("U_FixName", 0, "");
                oDS_PS_FX005H.SetValue("U_PostDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                oDS_PS_FX005H.SetValue("U_PostQty", 0, "0");
                oDS_PS_FX005H.SetValue("U_PostAmt", 0, "0");
                oDS_PS_FX005H.SetValue("U_StopYN", 0, "N");
                oDS_PS_FX005H.SetValue("U_LongYear", 0, "0");
                oDS_PS_FX005H.SetValue("U_DepRate", 0, "0");
                oDS_PS_FX005H.SetValue("U_CardCode", 0, "");
                oDS_PS_FX005H.SetValue("U_CardName", 0, "");
                oDS_PS_FX005H.SetValue("U_Kw", 0, "0");
                oDS_PS_FX005H.SetValue("U_Volt", 0, "0");
                oDS_PS_FX005H.SetValue("U_TeamCode", 0, "");
                oDS_PS_FX005H.SetValue("U_TeamNm", 0, "");
                oDS_PS_FX005H.SetValue("U_RspCode", 0, "");
                oDS_PS_FX005H.SetValue("U_RspNm", 0, "");
                oDS_PS_FX005H.SetValue("U_Place", 0, "");
                oDS_PS_FX005H.SetValue("U_Comments", 0, "");
                oDS_PS_FX005H.SetValue("U_TempChr1", 0, "");
                oDS_PS_FX005H.SetValue("U_PrcCode", 0, "");
                oDS_PS_FX005H.SetValue("U_PrcName", 0, "");

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
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
        /// PS_FX005_DisplayFixData
        /// </summary>
        private void PS_FX005_DisplayFixData(string DocEntry)
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (!string.IsNullOrEmpty(DocEntry.ToString().Trim()))
                {
                    sQry = "EXEC [PS_FX005_02] '" + DocEntry + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount == 0)
                    {
                        dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.:", "W");
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        PS_FX005_LoadCaption();
                        throw new Exception();
                    }

                    oDS_PS_FX005H.SetValue("U_Pic01", 0, "");
                    oDS_PS_FX005H.SetValue("U_Pic02", 0, "");
                    oDS_PS_FX005H.SetValue("U_Pic03", 0, "");
                    oDS_PS_FX005H.SetValue("U_Pic04", 0, "");
                    oDS_PS_FX005H.SetValue("U_Pic05", 0, "");
                    oDS_PS_FX005H.SetValue("U_Pic06", 0, "");

                    oDS_PS_FX005H.SetValue("U_Pic01", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_01.BMP");
                    oDS_PS_FX005H.SetValue("U_Pic02", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_02.BMP");
                    oDS_PS_FX005H.SetValue("U_Pic03", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_03.BMP");
                    oDS_PS_FX005H.SetValue("U_Pic04", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_04.BMP");
                    oDS_PS_FX005H.SetValue("U_Pic05", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_05.BMP");
                    oDS_PS_FX005H.SetValue("U_Pic06", 0, "\\\\191.1.1.220\\Asset_Pic\\" + oRecordSet.Fields.Item("FixCode").Value.ToString().Trim() + "_" + oRecordSet.Fields.Item("SubCode").Value.ToString().Trim() + "_06.BMP");

                    oDS_PS_FX005H.SetValue("DocEntry", 0, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_BPLId", 0, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_ClasCode", 0, oRecordSet.Fields.Item("ClasCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_SubDiv", 0, oRecordSet.Fields.Item("SubDiv").Value.ToString().Trim()); //Sub자산 여부(2015.07.06 송명규 추가)
                    oDS_PS_FX005H.SetValue("U_FixCode", 0, oRecordSet.Fields.Item("FixCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_SubCode", 0, oRecordSet.Fields.Item("SubCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_FixName", 0, oRecordSet.Fields.Item("FixName").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_PostDate", 0, oRecordSet.Fields.Item("PostDate").Value.ToString("yyyyMMdd").Trim());
                    oDS_PS_FX005H.SetValue("U_PostQty", 0, oRecordSet.Fields.Item("PostQty").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_PostAmt", 0, oRecordSet.Fields.Item("PostAmt").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_StopYN", 0, oRecordSet.Fields.Item("StopYN").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_LongYear", 0, oRecordSet.Fields.Item("LongYear").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_DepCode", 0, oRecordSet.Fields.Item("DepCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_DepRate", 0, oRecordSet.Fields.Item("DepRate").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_CardCode", 0, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_CardName", 0, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_Kw", 0, oRecordSet.Fields.Item("Kw").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_Volt", 0, oRecordSet.Fields.Item("Volt").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_TeamCode", 0, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_TeamNm", 0, oRecordSet.Fields.Item("TeamNm").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_RspCode", 0, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_RspNm", 0, oRecordSet.Fields.Item("RspNm").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_Place", 0, oRecordSet.Fields.Item("Place").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_Comments", 0, oRecordSet.Fields.Item("Comments").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_TempChr1", 0, oRecordSet.Fields.Item("TempChr1").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_PrcCode", 0, oRecordSet.Fields.Item("PrcCode").Value.ToString().Trim());
                    oDS_PS_FX005H.SetValue("U_PrcName", 0, oRecordSet.Fields.Item("PrcName").Value.ToString().Trim());

                }
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                PS_FX005_LoadCaption();
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
        /// PS_FX005_LoadPic
        /// </summary>
        private void PS_FX005_LoadPic(string pPictureControlName)
        {
            string errMessage = string.Empty;
            string sFilePath;
            string sFileName;
            string sFileExtension;
            string sQry;
            string SaveFolders;
            FileSystemObject FSO = new FileSystemObject();
            FileListBoxForm fileListBoxForm = new FileListBoxForm();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_FX005H.GetValue("U_FixCode", 0).ToString().Trim()))
                {
                    throw new Exception();
                }
                SaveFolders = "\\\\191.1.1.220\\Asset_Pic";

                //사진 불러오기
                sFilePath = fileListBoxForm.OpenDialog(fileListBoxForm, "*.BMP|*.BMP", "파일선택", "C:\\");
                if (string.IsNullOrEmpty(sFilePath))
                {
                    errMessage = "*.BMP|*.BMP 이미지가 선택되지 않았습니다.";
                    throw new Exception();
                }

                //확장자 확인
                sFileName = Path.GetFileName(sFilePath);
                sFileExtension = Path.GetExtension(sFilePath);
                if (sFileExtension.ToUpper() != ".BMP")
                {
                    errMessage = "BMP 확장자만 가능합니다.";
                    throw new Exception();
                }

                string imageFileName = null;
                if (pPictureControlName == "Pic01")
                {
                    imageFileName = "_01.BMP";
                }
                else if (pPictureControlName == "Pic02")
                {
                    imageFileName = "_02.BMP";
                }
                else if (pPictureControlName == "Pic03")
                {
                    imageFileName = "_03.BMP";
                }
                else if (pPictureControlName == "Pic04")
                {
                    imageFileName = "_04.BMP";
                }
                else if (pPictureControlName == "Pic05")
                {
                    imageFileName = "_05.BMP";
                }
                else if (pPictureControlName == "Pic06")
                {
                    imageFileName = "_06.BMP";
                }

                //서버에 기존 파일 체크
                FileInfo fileInfo = new FileInfo(SaveFolders + "\\" + sFileName);
                if (fileInfo.Exists)
                {
                    FSO.DeleteFile(SaveFolders + "\\" + sFileName);
                }
                FSO.CopyFile(sFilePath, SaveFolders + "\\" + oDS_PS_FX005H.GetValue("U_FixCode", 0).ToString().Trim() + "_" + oDS_PS_FX005H.GetValue("U_SubCode", 0).ToString().Trim() + imageFileName);

                sQry = " EXEC [PS_FX005_03] '";
                sQry += pPictureControlName + "','";
                sQry += oDS_PS_FX005H.GetValue("U_FixCode", 0).ToString().Trim() + "', '";
                sQry += oDS_PS_FX005H.GetValue("U_SubCode", 0).ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.MessageBox("사진이 업로드 되었습니다.");
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
            string sQry;
            string pFixCode;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_save")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_FX005_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_FX005_AddData(pVal) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            pFixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();
                            sQry = "EXEC [PS_Table_history] '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "','FX005','" + PSH_Globals.oCompany.UserSignature + "'";
                            oRecordSet.DoQuery(sQry);

                            PS_FX005_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_FX005_LoadCaption();
                            PS_FX005_MTX01(pFixCode);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_FX005_UpdateData(pVal) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            pFixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();
                            sQry = "EXEC [PS_Table_history] '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "','FX005','" + PSH_Globals.oCompany.UserSignature + "'";
                            oRecordSet.DoQuery(sQry);

                            PS_FX005_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PS_FX005_LoadCaption();
                            PS_FX005_MTX01(pFixCode);
                        }
                    }
                    else if (pVal.ItemUID == "Btn_ret")
                    {
                        PS_FX005_MTX01(oForm.Items.Item("SFixCode").Specific.Value.ToString().Trim());
                    }
                    else if (pVal.ItemUID == "Btn_del")
                    {
                        pFixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();
                        PS_FX005_DeleteData();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    }
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
                        if (pVal.ItemUID == "CardCde")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "TeamCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("TeamCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "RspCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("RspCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "PrcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("PrcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "FixCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("FixCode").Specific.Value))
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim()))
                                {
                                    dataHelpClass.MDC_GF_Message("분류가 선택되지 않았습니다. 확인하세요.", "W");
                                }
                                else
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            string BPLId;
            string ClasCode;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "ClasCode" || pVal.ItemUID == "BPLId")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (oForm.Items.Item("SubDiv").Specific.Value.ToString().Trim() == "N")
                            {
                                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim();
                                sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_FixCode),3)) + 1)),3) From [@PS_FX005H] Where U_BPLId = '" + BPLId + "' And U_ClasCode = '" + ClasCode + "'";
                                oRecordSet.DoQuery(sQry);

                                if (!string.IsNullOrEmpty(BPLId.ToString().Trim()) && ClasCode.ToString().Trim() != "%")
                                {
                                    if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
                                    {
                                        oForm.Items.Item("FixCode").Specific.String = BPLId + ClasCode + "001";
                                    }
                                    else
                                    {
                                        oForm.Items.Item("FixCode").Specific.String = BPLId + ClasCode + oRecordSet.Fields.Item(0).Value.ToString().Trim();
                                    }
                                    oForm.Items.Item("SubCode").Specific.Value = "001";
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "SubDiv")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (oForm.Items.Item("SubDiv").Specific.Value.ToString().Trim() == "Y")
                            {
                                oForm.Items.Item("FixCode").Enabled = true;
                                oForm.Items.Item("FixCode").Specific.Value = "";
                            }
                            else
                            {
                                oForm.Items.Item("FixCode").Enabled = false;
                                oForm.Items.Item("SubCode").Specific.Value = "001";
                                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                                ClasCode = oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim();

                                sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_FixCode),3)) + 1)),3) From [@PS_FX005H] Where U_BPLId = '" + BPLId + "' And U_ClasCode = '" + ClasCode + "'";
                                oRecordSet.DoQuery(sQry);

                                if (!string.IsNullOrEmpty(BPLId.ToString().Trim()) && ClasCode.ToString().Trim() != "%")
                                {
                                    if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
                                    {
                                        oForm.Items.Item("FixCode").Specific.String = BPLId + ClasCode + "001";
                                    }
                                    else
                                    {
                                        oForm.Items.Item("FixCode").Specific.String = BPLId + ClasCode + oRecordSet.Fields.Item(0).Value.ToString().Trim();
                                    }
                                    oForm.Items.Item("SubCode").Specific.Value = "001";
                                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            PS_FX005_DisplayFixData(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "ClasCode")
                    {
                        dataHelpClass.Set_ComboList(oForm.Items.Item("LongYear").Specific, "", "", true, false);
                        dataHelpClass.Set_ComboList(oForm.Items.Item("LongYear").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'FX007' and b.U_RelCd = '" + oForm.Items.Item("ClasCode").Specific.Value.ToString().Trim() + "' order by U_Seq", "", false, false);
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
            try
            {
                if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            PS_FX005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "TeamCode")
                        {
                            PS_FX005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "RspCode")
                        {
                            PS_FX005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "PrcCode")
                        {
                            PS_FX005_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        if (pVal.ItemUID == "FixCode")
                        {
                            PS_FX005_FlushToItemValue(pVal.ItemUID, 0, "");
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX005H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX005L);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Pic01" || pVal.ItemUID == "Pic02" || pVal.ItemUID == "Pic03" || pVal.ItemUID == "Pic04" || pVal.ItemUID == "Pic05" || pVal.ItemUID == "Pic06")
                    {
                        PS_FX005_LoadPic(pVal.ItemUID);
                        PS_FX005_DisplayFixData(oForm.Items.Item("DocEntry").Specific.Value);
                        BubbleEvent = false;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            PS_FX005_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            BubbleEvent = false;
                            PS_FX005_LoadCaption();
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
