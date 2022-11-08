using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정산공제대상자정보등록
    /// </summary>
    internal class PH_PY403 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat;
        private SAPbouiCOM.DBDataSource oDS_PH_PY403A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY403B;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY403.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY403_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY403");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";
                
                oForm.Freeze(true);
                PH_PY403_CreateItems();
                PH_PY403_ComboBox_Setting();
                PH_PY403_EnableMenus();
                PH_PY403_FormItemEnabled();
                PH_PY403_AddMatrixRow();
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
        /// <returns></returns>
        private void PH_PY403_CreateItems()
        {
            try
            {
                oDS_PH_PY403A = oForm.DataSources.DBDataSources.Item("@PH_PY403A");
                oDS_PH_PY403B = oForm.DataSources.DBDataSources.Item("@PH_PY403B");
                oMat = oForm.Items.Item("Mat01").Specific;
                oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat.AutoResizeColumns();

                // 정산년도
                oForm.Items.Item("YY").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);

                //성명
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                // 사번
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// ComboBox_Setting
        /// </summary>
        private void PH_PY403_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                //주택구분
                oForm.Items.Item("House").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("House").Specific.ValidValues.Add("0", "무주택");
                oForm.Items.Item("House").Specific.ValidValues.Add("1", "1주택");
                oForm.Items.Item("House").Specific.ValidValues.Add("2", "2주택이상");
                oForm.Items.Item("House").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("House").DisplayDesc = true;

                //세대구분
                oForm.Items.Item("Saede").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Saede").Specific.ValidValues.Add("Y", "세대주");
                oForm.Items.Item("Saede").Specific.ValidValues.Add("N", "세대원");
                oForm.Items.Item("Saede").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Saede").DisplayDesc = true;

                //부녀자공제
                oForm.Items.Item("Woman").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Woman").Specific.ValidValues.Add("Y", "해당");
                oForm.Items.Item("Woman").Specific.ValidValues.Add("N", "비해당");
                oForm.Items.Item("Woman").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Woman").DisplayDesc = true;

                //한무모공제
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("Y", "해당");
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("N", "비해당");
                oForm.Items.Item("Sparent").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Sparent").DisplayDesc = true;

                //재직구분
                oForm.DataSources.UserDataSources.Add("Status_1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT statusID, name FROM [OHST] Order By 1 ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Status_1").Specific, "N");


                //관계 oMat
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oMat.Columns.Item("Relate").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oMat.Columns.Item("Relate").DisplayDesc = true;

                //소득유무 oMat
                oMat.Columns.Item("Soduk").ValidValues.Add("", "선택");
                oMat.Columns.Item("Soduk").ValidValues.Add("Y", "소득있음");
                oMat.Columns.Item("Soduk").ValidValues.Add("N", "소득없음");
                oMat.Columns.Item("Soduk").DisplayDesc = true;

                //장애인코드 oMat
                oMat.Columns.Item("HDCode").ValidValues.Add("", "해당업음");
                oMat.Columns.Item("HDCode").ValidValues.Add("1", "장애인복지법에 따른 장애인");
                oMat.Columns.Item("HDCode").ValidValues.Add("2", "국가유공자등 예우및지원에 관한 법률에 따른 상이자 및 이와 유사한자로서 근로능력이없는자");
                oMat.Columns.Item("HDCode").ValidValues.Add("3", "그 밖에 항시 치료를 요하는 중증환자");
                oMat.Columns.Item("HDCode").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY403_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY403_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    PH_PY403_FormClear();            //폼 DocEntry 세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true);  //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true);  //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY403_DataValidCheck()
        {
            bool returnValue = false;

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //사번
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //대출금액
                if (oDS_PH_PY403A.GetValue("U_House", 0).ToString().Trim() == "")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("주택구분은 필수입니다. 입력하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("House").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환기간
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_RpmtPrd", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환기간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtPrd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //라인
                if (oMat.VisualRowCount > 1)
                {
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return returnValue;
                }

                oMat.FlushToDataSource();
                if (oDS_PH_PY403B.Size > 1)
                {
                    oDS_PH_PY403B.RemoveRecord(oDS_PH_PY403B.Size - 1);
                }
                oMat.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            return returnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <param name="prmRow"></param>
        /// <returns></returns>
        private bool PH_PY403_Validate(string ValidateType, int prmRow)
        {
            bool returnValue = false;
            int ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY403A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    ErrNumm = 1;
                    throw new Exception();
                }

                if (ValidateType == "수정")
                {

                    if (oDS_PH_PY403B.GetValue("U_RpmtYN", prmRow - 1) == "Y")
                    {
                        PSH_Globals.SBO_Application.MessageBox("상환이 완료된 행입니다. 수정할 수 없습니다.");
                        throw new Exception();
                    }
                }
                else if (ValidateType == "행삭제")
                {
                    if (oDS_PH_PY403B.GetValue("U_RpmtYN", oLastColRow - 1) == "Y")
                    {
                        PSH_Globals.SBO_Application.MessageBox("상환이 완료된 행입니다. 수정할 수 없습니다.");
                        throw new Exception();
                    }
                }
                else if (ValidateType == "취소")
                {
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY403_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY403'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
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
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY403_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat.FlushToDataSource();
                oRow = oMat.VisualRowCount;

                if (oMat.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY403B.Size <= oMat.VisualRowCount)
                        {
                            oDS_PH_PY403B.InsertRecord(oRow);
                        }
                        oDS_PH_PY403B.Offset = oRow;
                        oDS_PH_PY403B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY403B.SetValue("U_Chk", oRow, "N");
                        oDS_PH_PY403B.SetValue("U_KName", oRow, "");
                        oDS_PH_PY403B.SetValue("U_Relate", oRow, "");
                        oDS_PH_PY403B.SetValue("U_JuminNo", oRow, "");
                        oDS_PH_PY403B.SetValue("U_Soduk", oRow, "");
                        oDS_PH_PY403B.SetValue("U_HDCode", oRow, "");
                        oMat.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY403B.Offset = oRow - 1;
                        oDS_PH_PY403B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY403B.SetValue("U_Chk", oRow - 1, "N");
                        oDS_PH_PY403B.SetValue("U_KName", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_Relate", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_JuminNo", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_Soduk", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_HDCode", oRow - 1, "");
                        oMat.LoadFromDataSource();
                    }
                }
                else if (oMat.VisualRowCount == 0)
                {
                    oDS_PH_PY403B.Offset = oRow;
                    oDS_PH_PY403B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY403B.SetValue("U_Chk", oRow, "N");
                    oDS_PH_PY403B.SetValue("U_KName", oRow, "");
                    oDS_PH_PY403B.SetValue("U_Relate", oRow, "");
                    oDS_PH_PY403B.SetValue("U_JuminNo", oRow, "");
                    oDS_PH_PY403B.SetValue("U_Soduk", oRow, "");
                    oDS_PH_PY403B.SetValue("U_HDCode", oRow, "");
                    oMat.LoadFromDataSource();
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
        /// PH_PY403_MTX01
        /// </summary>
        private void PH_PY403_MTX01()
        {
            int i;
            string sQry;
            string errCode = string.Empty;

            string Param01;
            string Param02;
            string Param03;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("LoanAmt").Specific.Value;
                Param02 = oForm.Items.Item("LoanDate").Specific.Value;
                Param03 = oForm.Items.Item("RpmtPrd").Specific.Value;

                sQry = "EXEC PH_PY403_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                oMat.Clear();
                oMat.FlushToDataSource();
                oMat.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    oMat.Clear();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY403B.InsertRecord(i);
                    }

                    //마지막 빈행 추가
                    if (i == oRecordSet.RecordCount)
                    {
                        oDS_PH_PY403B.Offset = i;
                        oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PH_PY403B.SetValue("U_Cnt", i, "");
                        oDS_PH_PY403B.SetValue("U_PayDate", i, "");
                        oDS_PH_PY403B.SetValue("U_RpmtAmt", i, "0");
                        oDS_PH_PY403B.SetValue("U_TotRpmt", i, "0");
                    }
                    else
                    {
                        oDS_PH_PY403B.Offset = i;
                        oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PH_PY403B.SetValue("U_Cnt", i, oRecordSet.Fields.Item("Cnt").Value);
                        oDS_PH_PY403B.SetValue("U_PayDate", i, oRecordSet.Fields.Item("PayDate").Value);
                        oDS_PH_PY403B.SetValue("U_RpmtAmt", i, oRecordSet.Fields.Item("RpmtAmt").Value);
                        oDS_PH_PY403B.SetValue("U_TotRpmt", i, oRecordSet.Fields.Item("TotRpmt").Value);

                        oRecordSet.MoveNext();
                    }

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }

                oMat.LoadFromDataSource();
                oMat.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY403_CalDataCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY403_CalDataCheck()
        {
            bool returnValue = false;
            
            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //사번
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //대출일자
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_LoanDate", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("대출일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("LoanDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //대출금액
                if (Convert.ToDouble(oDS_PH_PY403A.GetValue("U_LoanAmt", 0).ToString().Trim()) == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("대출금액은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("LoanAmt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //이자율
                if (oDS_PH_PY403A.GetValue("U_IntRate", 0).ToString().Trim() == "0.0")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("이자율은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("IntRate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                //상환기간
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_RpmtPrd", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환기간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtPrd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY403_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY403_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
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
                                PH_PY403_FormItemEnabled();
                                PH_PY403_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY403_FormItemEnabled();
                                PH_PY403_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY403_FormItemEnabled();
                            }
                        }
                    }
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
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "CntcName" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcName").Specific.Value.ToString().Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string CLTCOD;
            string CntcCode;
            string CntcName;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CntcCode":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

                                sQry = "  Select Code,";
                                sQry += "        FullName = U_FullName,";
                                sQry += "        TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                            WHERE Code = '1'";
                                sQry += "                              And U_Code = U_TeamCode),''),";
                                sQry += "        RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                            WHERE Code = '2'";
                                sQry += "                              And U_Code = U_RspCode),''),";
                                sQry += "        Status = U_Status";
                                sQry += "   From [@PH_PY001A]";
                                sQry += "  Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += "    and Code = '" + CntcCode + "'";
                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("CntcName").Value = oRecordSet.Fields.Item("FullName").Value.ToString().Trim();
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("Status_1").Specific.Select(oRecordSet.Fields.Item("Status").Value.ToString().Trim());
                                break;

                            case "CntcName":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();

                                sQry = "Select Code,";
                                sQry += "      FullName = U_FullName,";
                                sQry += "      TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                           From [@PS_HR200L]";
                                sQry += "                          WHERE Code = '1'";
                                sQry += "                            And U_Code = U_TeamCode),''),";
                                sQry += "      RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                           From [@PS_HR200L]";
                                sQry += "                          WHERE Code = '2'";
                                sQry += "                            And U_Code = U_RspCode),''),";
                                sQry += "      Status = U_Status";
                                sQry += " From [@PH_PY001A]";
                                sQry += "Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += "  And U_status <> '5'"; // 퇴사자 제외
                                sQry += "  and U_FullName = '" + CntcName + "'";
                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("CntcCode").Value = oRecordSet.Fields.Item("Code").Value.ToString().Trim();
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("Status_1").Specific.Select(oRecordSet.Fields.Item("Status").Value.ToString().Trim());
                                break;

                            case "Mat01":

                                if (pVal.ColUID == "KName")
                                {
                                    oMat.FlushToDataSource();
                                    oDS_PH_PY403B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                    oMat.LoadFromDataSource();

                                    if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PH_PY403_AddMatrixRow();
                                    }
                                }

                                oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat.AutoResizeColumns();
                                break;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat.LoadFromDataSource();

                    PH_PY403_FormItemEnabled();
                    PH_PY403_AddMatrixRow();
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY403A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY403B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
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
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            if (PH_PY403_Validate("행삭제", 0) == false)
                            {
                                BubbleEvent = false;
                                oForm.Freeze(false);
                                return;
                            }
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY403A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY403_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            if (oMat.RowCount != oMat.VisualRowCount)
                            {
                                oMat.FlushToDataSource();

                                while (i <= oDS_PH_PY403B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY403B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY403B.Size; i++)
                                {
                                    oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat.LoadFromDataSource();
                            }
                            break;
                        case "1287": //복제

                            oForm.Freeze(true);
                            oDS_PH_PY403A.SetValue("DocEntry", 0, "");

                            for (i = 0; i <= oMat.VisualRowCount - 1; i++)
                            {
                                oMat.FlushToDataSource();
                                oDS_PH_PY403B.SetValue("DocEntry", i, "");
                                oDS_PH_PY403B.SetValue("U_PayYN", i, "N");
                                oMat.LoadFromDataSource();
                            }

                            oForm.Freeze(false);
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
        }
    }
}

