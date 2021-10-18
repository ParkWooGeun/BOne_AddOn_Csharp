using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사업장정보등록
    /// </summary>
    internal class PH_PY005 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.DBDataSource oDS_PH_PY005A;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {   
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY005.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PH_PY005_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY005");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                oForm.Items.Item("FLD01").Specific.Select();
                oForm.Visible = true;
                PH_PY005_CreateItems();
                PH_PY005_EnableMenus();
                PH_PY005_SetDocument(oFormDocEntry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY005_CreateItems()
        {
            string sQry;
            int i;

            SAPbouiCOM.CheckBox oCheck = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            oForm.Freeze(true);

            try
            {
                oDS_PH_PY005A = oForm.DataSources.DBDataSources.Item("@PH_PY005A");

                //제출인구분
                oCombo = oForm.Items.Item("TaxDGbn").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P006' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");
                oForm.Items.Item("TaxDGbn").DisplayDesc = true;

                // 납세자구분
                oCombo = oForm.Items.Item("BUSTYP").Specific;
                oCombo.ValidValues.Add("1", "개인");
                oCombo.ValidValues.Add("8", "법인");
                oForm.Items.Item("BUSTYP").DisplayDesc = true;

                // 사업자단위과세
                oCombo = oForm.Items.Item("SAUPJA").Specific;
                oCombo.ValidValues.Add("N", "N");
                oCombo.ValidValues.Add("Y", "Y");
                oForm.Items.Item("SAUPJA").DisplayDesc = true;

                // 원천신고구분
                oCombo = oForm.Items.Item("SINTYP").Specific;
                oCombo.ValidValues.Add("1", "매월");
                oCombo.ValidValues.Add("2", "반기");
                oForm.Items.Item("SINTYP").DisplayDesc = true;

                // 일괄납부여부
                oCombo = oForm.Items.Item("ILGTYP").Specific;
                oCombo.ValidValues.Add("1", "부");
                oCombo.ValidValues.Add("2", "여");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ILGTYP").DisplayDesc = true;

                //자동체번 생성규칙
                oCombo = oForm.Items.Item("WCHCLT").Specific;
                sQry = "SELECT BPLId, BPLName FROM [OBPL]";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");

                oForm.Items.Item("WCHCLT").DisplayDesc = true;

                //자동체번 생성규칙
                oCombo = oForm.Items.Item("SUPCLT").Specific;
                sQry = "SELECT BPLId, BPLName FROM [OBPL]";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");
                oForm.Items.Item("SUPCLT").DisplayDesc = true;

                //자동체번 생성규칙
                oCombo = oForm.Items.Item("JUMCLT").Specific;
                sQry = "SELECT BPLId, BPLName FROM [OBPL]";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");
                oForm.Items.Item("JUMCLT").DisplayDesc = true;

                //------------------------------------------------------------------------------------------------
                //세부정보
                //------------------------------------------------------------------------------------------------
                //사원번호구성체계
                oCombo = oForm.Items.Item("AutoChk").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P005' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "");
                oForm.Items.Item("AutoChk").DisplayDesc = true;

                // 이행상황신고서집계방법
                oCombo = oForm.Items.Item("WCHTYP").Specific;
                oCombo.ValidValues.Add("1", "귀속연월");
                oCombo.ValidValues.Add("2", "지급연월");
                oCombo.ValidValues.Add("3", "귀속연월 OR 지급연월");
                oCombo.ValidValues.Add("4", "귀속연월 AND 지급연월");
                oCombo.ValidValues.Add("5", "신고연월");
                oForm.Items.Item("WCHTYP").DisplayDesc = true;

                // 사원번호 자릿수
                oCombo = oForm.Items.Item("EmpTLen").Specific;
                for (i = 1; i <= 8; i++)
                {
                    oCombo.ValidValues.Add(Convert.ToString(i), i + " Length");
                }
                // 결재란수(기본 6개)
                oCombo = oForm.Items.Item("EmpType").Specific;
                for (i = 1; i <= 6; i++)
                {
                    oCombo.ValidValues.Add(Convert.ToString(i), Convert.ToString(i));
                }

                // Check 버튼
                oCheck = oForm.Items.Item("govIDChk").Specific;
                oCheck.ValOff = "N";
                oCheck.ValOn = "Y";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY005_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);

                if (oForm.Visible == false)
                {
                    oForm.Visible = true;
                }

                oForm.Update();
                //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCheck);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY005_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY005_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PH_PY005_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PH_PY005_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY005_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY005_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY005_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    //Call CLTCOD_Select(oForm, "CLTCOD", False)
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY005_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY005_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY005'", "");
                if (Convert.ToInt32(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY005_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY005_DataValidCheck()
        {
            bool returnValue = false;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.Value_ChkYn("[@PH_PY005A]", "Code", "'" + oForm.Items.Item("CLTCode").Specific.Value + "'", "") == false)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("이미 저장되어져 있는 코드가 존재합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        oForm.Items.Item("CLTCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return returnValue;
                    }
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_BusNum", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업자번호는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BusNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }
                else if (dataHelpClass.Value_ChkYn("[@PH_PY005A]", "U_BusNum", "'" + oDS_PH_PY005A.GetValue("U_BusNum", 0).Trim() + "'", " AND Code <> '" + oForm.Items.Item("Code").Specific.Value + "'") == false)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업자번호가 중복되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BusNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_BUSTYP", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("납세자구분은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BUSTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_SINTYP", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("원천신고구분은 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("SINTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_WCHCLT", 0).Trim()))
                {
                    oDS_PH_PY005A.SetValue("U_WCHCLT", 0, oForm.Items.Item("Code").Specific.Value);
                    oForm.Items.Item("WCHCLT").Update();
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_SUPCLT", 0).Trim()))
                {
                    oDS_PH_PY005A.SetValue("U_SUPCLT", 0, oForm.Items.Item("Code").Specific.Value);
                    oForm.Items.Item("SUPCLT").Update();
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_JUMCLT", 0).Trim()))
                {
                    oDS_PH_PY005A.SetValue("U_JUMCLT", 0, oForm.Items.Item("Code").Specific.Value);
                    oForm.Items.Item("JUMCLT").Update();
                }

                //oForm.PaneLevel = 2;
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_AutoChk", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사원번호 구성 체계는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("AutoChk").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_EmpTLen", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사원번호 자릿수는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("EmpTLen").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY005A.GetValue("U_EmpType", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결재란 수는 필수입니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("EmpType").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return returnValue;
                }

                // Code & Name 생성
                oDS_PH_PY005A.SetValue("Code", 0, oDS_PH_PY005A.GetValue("U_CLTCode", 0).Trim());
                oDS_PH_PY005A.SetValue("NAME", 0, oDS_PH_PY005A.GetValue("U_CLTName", 0).Trim());

                returnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY005_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY005_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errCode = string.Empty;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY005A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (ValidateType == "수정")
                {

                }
                else if (ValidateType == "행삭제")
                {

                }
                else if (ValidateType == "취소")
                {

                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else 
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            
            return returnValue;
        }

        /// <summary>
        /// 결재란 수 Enable 조정
        /// </summary>
        private void Sign_Enabled()
        {
            short i;
            short MAXCNT;

            MAXCNT = Convert.ToInt16(oDS_PH_PY005A.GetValue("U_EmpType", 0));

            for (i = 1; i <= 8; i++)
            {
                if (MAXCNT >= i)
                {
                    oForm.Items.Item("Sign0" + i).Enabled = true;
                }
                else
                {
                    oDS_PH_PY005A.SetValue("U_Sign0" + i, 0, "");
                    oForm.Items.Item("Sign0" + i).Enabled = false;
                }
            }

        }

        /// <summary>
        /// 직인 이미지 등록(사용안됨)
        /// </summary>
        private void Picture_Save()
        {
            //double L_Code = 0;
            //string SEALIMG = null;

            //L_Code = Conversion.Val(oForm.Items.Item("Code").Specific.String);

            //if (Conversion.Val(Convert.ToString(L_Code)) == 0)
            //{
            //    return;
            //}

            //L_Code = L_Code * -1;

            //if (oForm.Items.Item("IMGCHK").Specific.Checked == false)
            //{

            //    PSH_Globals.SBO_Application.StatusBar.SetText("세무서식에 사용할 직인 등록을 체크하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return;

            //}

            ////UPGRADE_WARNING: FileListBoxForm.OpenDialog() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            //SEALIMG = My.MyProject.Forms.FileListBoxForm.OpenDialog(ref oForm, ref "graphic Files (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF", ref "파일선택", ref "C:\\");

            //if (string.IsNullOrEmpty(Strings.Trim(SEALIMG)))
            //    return;

            //string sQry = null;

            //int iFileLen = 0;
            //int iDataFile = 0;
            //short i = 0;
            //short iFrag = 0;
            //short iChunks = 0;
            //byte[] iChunk = null;

            //const short iChunkSize = 16348;

            //sQry = "SELECT EmpID, FILLEN, FILIMG FROM MDC_PAYPIC WHERE EmpID = " + L_Code + "";
            //MDC_Globals.g_ADORS1 = new ADODB.Recordset();
            //MDC_Globals.g_ADORS1.Open(sQry, MDC_Globals.g_ERPDMS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic);

            //iDataFile = FreeFile();

            //FileSystem.FileOpen(iDataFile, SEALIMG, OpenMode.Binary, OpenAccess.Read);

            //iFileLen = FileSystem.LOF(iDataFile);

            //if (iFileLen > 0)
            //{

            //    if (MDC_Globals.g_ADORS1.EOF)
            //        MDC_Globals.g_ADORS1.AddNew();
            //    MDC_Globals.g_ADORS1.Fields["EmpID"].Value = L_Code;
            //    MDC_Globals.g_ADORS1.Fields["FILLEN"].Value = iFileLen;
            //    /// 길이
            //    iChunks = iFileLen / iChunkSize;
            //    iFrag = iFileLen % iChunkSize;
            //    //UPGRADE_WARNING: Null/IsNull() 사용이 감지되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            //    MDC_Globals.g_ADORS1.Fields["FILIMG"].AppendChunk(System.DBNull.Value);
            //    iChunk = new byte[iFrag + 1];
            //    //UPGRADE_WARNING: Get이(가) FileGet(으)로 업그레이드되어 새 동작을 가집니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            //    FileSystem.FileGet(iDataFile, iChunk);
            //    MDC_Globals.g_ADORS1.Fields["FILIMG"].AppendChunk(iChunk);
            //    iChunk = new byte[iChunkSize + 1];
            //    for (i = 1; i <= iChunks; i++)
            //    {
            //        //UPGRADE_WARNING: Get이(가) FileGet(으)로 업그레이드되어 새 동작을 가집니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            //        FileSystem.FileGet(iDataFile, iChunk);
            //        MDC_Globals.g_ADORS1.Fields["FILIMG"].AppendChunk(iChunk);
            //    }

            //}

            //FileSystem.FileClose(iDataFile);

            //MDC_Globals.g_ADORS1.Update();

            ////UPGRADE_NOTE: g_ADORS1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            //MDC_Globals.g_ADORS1 = null;

            //return;
            //Error_Message:
            ///// Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
            ////UPGRADE_NOTE: g_ADORS1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            //MDC_Globals.g_ADORS1 = null;
            //MDC_Globals.Sbo_Application.StatusBar.SetText("Picture_Save Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

        }

        /// <summary>
        /// 직인 이미지 삭제
        /// </summary>
        private void Picture_Delete()
        {
            //string sQry = null;
            //double L_Code = 0;
            //SAPbobsCOM.Recordset oRecordSet = null;

            //L_Code = Conversion.Val(oForm.Items.Item("Code").Specific.String);

            //if (Conversion.Val(Convert.ToString(L_Code)) == 0)
            //    return;

            //oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //L_Code = L_Code * -1;

            //sQry = "DELETE FROM [MDC_PAYPIC] WHERE EmpID = " + L_Code + "";
            //oRecordSet.DoQuery(sQry);

            ////UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            //oRecordSet = null;
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
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY005_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CBtn1")
                    {
                        oForm.Items.Item("TaxCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.ItemUID == "CBtn2")
                    {
                        oForm.Items.Item("EmpID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.ItemUID == "CBtn6")
                    {
                        oForm.Items.Item("BNKCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                        return;
                    }
                    else if (pVal.ItemUID == "FLD01")
                    {
                        oForm.PaneLevel = 1;
                    }
                    else if (pVal.ItemUID == "FLD02")
                    {
                        oForm.PaneLevel = 2;
                    }
                    else if (pVal.ItemUID == "INSBtn1")
                    {
                        Picture_Save();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ActionSuccess == true)
                    {
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY005_FormItemEnabled();

                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY005_FormItemEnabled();

                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY005_FormItemEnabled();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemUID == "Code" && pVal.CharPressed == 9 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        if (dataHelpClass.Value_ChkYn("[@PH_PY005A]", "Code", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "") == false)
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("동일한 자사코드가 존재합니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                    }

                    if (pVal.ItemUID == "TaxCode" && pVal.CharPressed == 9)
                    {
                        if (dataHelpClass.Value_ChkYn("[@PS_HR200L]", "U_Code", "'" + oForm.Items.Item(pVal.ItemUID).Specific.String + "'", " AND Code='P007'") == true)
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                            return;
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
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
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        switch (pVal.ItemUID)
                        {
                            case "TaxDGbn": // 대리인구분
                                if (oForm.Items.Item(pVal.ItemUID).Specific.Selected == null)
                                {
                                    oDS_PH_PY005A.SetValue("U_TaxDGbn", 0, "");
                                    oDS_PH_PY005A.SetValue("U_TaxDGnm", 0, "");
                                }
                                else
                                {
                                    oDS_PH_PY005A.SetValue("U_TaxDGbn", 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                    oDS_PH_PY005A.SetValue("U_TaxDGnm", 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Description);

                                    if (oDS_PH_PY005A.GetValue("U_TaxDGbn", 0) == "2")
                                    {
                                        oDS_PH_PY005A.SetValue("U_TaxDNam", 0, oDS_PH_PY005A.GetValue("Code", 0));
                                        oDS_PH_PY005A.SetValue("U_TaxDBus", 0, oDS_PH_PY005A.GetValue("U_BusNum", 0));
                                        oForm.Items.Item("TaxDNam").Update();
                                        oForm.Items.Item("TaxDBus").Update();
                                    }
                                    oForm.Update();
                                }
                                break;
                            case "AutoChk":
                                if (oForm.Items.Item(pVal.ItemUID).Specific.Selected != null)
                                {
                                    oDS_PH_PY005A.SetValue("U_AutoChk", 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);

                                    switch (oDS_PH_PY005A.GetValue("U_AutoChk", 0).Trim())
                                    {
                                        case "1":
                                        case "2":
                                        case "3":
                                            oDS_PH_PY005A.SetValue("U_EmpTLen", 0, "8");
                                            break;
                                        case "4":
                                        case "5":
                                        case "6":
                                            oDS_PH_PY005A.SetValue("U_EmpTLen", 0, "6");
                                            break;
                                        case "7":
                                            oDS_PH_PY005A.SetValue("U_EmpTLen", 0, "5");
                                            break;
                                    }
                                    oForm.Items.Item("EmpTLen").Update();
                                }
                                break;
                            case "EmpType":
                                Sign_Enabled();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "Code":
                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                                {
                                    oDS_PH_PY005A.SetValue("U_CLTCode", 0, "");
                                    oDS_PH_PY005A.SetValue("Code", 0, "");
                                }
                                else
                                {
                                    oDS_PH_PY005A.SetValue("U_CLTCode", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToUpper());
                                    oDS_PH_PY005A.SetValue("Code", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToUpper());
                                }
                                break;
                            case "Name":
                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                                {
                                    oDS_PH_PY005A.SetValue("Name", 0, "");
                                    oDS_PH_PY005A.SetValue("U_CLTName", 0, "");
                                }
                                else
                                {
                                    oDS_PH_PY005A.SetValue("Name", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToUpper());
                                    oDS_PH_PY005A.SetValue("U_CLTName", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToUpper());
                                }
                                oForm.Items.Item("CLTName").Update();
                                break;
                            case "EmpID":
                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                                {
                                    oDS_PH_PY005A.SetValue("U_EmpID", 0, "");
                                    oDS_PH_PY005A.SetValue("U_ComPrt", 0, "");
                                }
                                else
                                {
                                    oDS_PH_PY005A.SetValue("U_EmpID", 0, oForm.Items.Item(pVal.ItemUID).Specific.String);
                                    oDS_PH_PY005A.SetValue("U_ComPrt", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                                }
                                oForm.Items.Item("ComPrt").Update();
                                break;
                            case "TaxCode":
                                if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                                {
                                    oDS_PH_PY005A.SetValue("U_TaxCode", 0, "");
                                    oDS_PH_PY005A.SetValue("U_TaxName", 0, "");
                                    oDS_PH_PY005A.SetValue("U_TaxAcct", 0, "");
                                }
                                else
                                {
                                    oDS_PH_PY005A.SetValue("U_TaxCode", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                    oDS_PH_PY005A.SetValue("U_TaxName", 0, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "And Code='P007'"));
                                    oDS_PH_PY005A.SetValue("U_TaxAcct", 0, dataHelpClass.Get_ReData("U_Char1", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", "And Code='P007'"));
                                }
                                oForm.Items.Item("TaxName").Update();
                                oForm.Items.Item("TaxAcct").Update();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY005A);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY005_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY005_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY005_FormItemEnabled();
                            oDS_PH_PY005A.SetValue("U_WCHTYP", 0, "5");
                            oDS_PH_PY005A.SetValue("U_SAUPJA", 0, "N");
                            oForm.Items.Item("WCHTYP").Update();
                            oForm.Items.Item("FLD01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY005_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                switch (pVal.ItemUID) {
                    case "Mat1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
