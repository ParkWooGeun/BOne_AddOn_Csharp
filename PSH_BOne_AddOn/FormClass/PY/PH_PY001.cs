using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사원마스터등록
    /// </summary>
    internal class PH_PY001 : PSH_BaseClass
    {
        #region 변수선언
        private string oFormUniqueID;

        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.Matrix oMat2;
        private SAPbouiCOM.Matrix oMat3;
        private SAPbouiCOM.Matrix oMat4;
        private SAPbouiCOM.Matrix oMat5;
        private SAPbouiCOM.Matrix oMat6;
        private SAPbouiCOM.Matrix oMat7;
        private SAPbouiCOM.Matrix oMat8;
        private SAPbouiCOM.Matrix oMat9;
        private SAPbouiCOM.Matrix oMat10;
        private SAPbouiCOM.Matrix oMat11;
        private SAPbouiCOM.Matrix oMat12;
        private SAPbouiCOM.Matrix oMat13;
        private SAPbouiCOM.Matrix oMat14;
        private SAPbouiCOM.Matrix oMat15;

        private SAPbouiCOM.DBDataSource oDS_PH_PY001A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001B;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001C;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001D;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001E;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001F;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001G;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001H;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001J;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001K;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001L;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001M;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001N;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001P;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001Q;
        private SAPbouiCOM.DBDataSource oDS_PH_PY001R;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        #endregion

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY001.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int loopCount = 1; loopCount <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); loopCount++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[loopCount - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[loopCount - 1].nodeValue = 16;
                }
                oFormUniqueID = "PH_PY001_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY001");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                oForm.Items.Item("FLD01").Specific.Select();
                oForm.Visible = true;
                PH_PY001_CreateItems();
                PH_PY001_EnableMenus();
                PH_PY001_SetDocument(oFormDocEntry01);
                PSH_Globals.ExecuteEventFilter(typeof(PH_PY001));
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
        private void PH_PY001_CreateItems()
        {
            string sQry;
            int i;

            SAPbouiCOM.CheckBox oCheck = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.Column oColumn = null;
            SAPbouiCOM.OptionBtn optBtn = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oForm.Freeze(true);

            try
            {
                oDS_PH_PY001A = oForm.DataSources.DBDataSources.Item("@PH_PY001A");
                oDS_PH_PY001B = oForm.DataSources.DBDataSources.Item("@PH_PY001B");
                oDS_PH_PY001C = oForm.DataSources.DBDataSources.Item("@PH_PY001C");
                oDS_PH_PY001D = oForm.DataSources.DBDataSources.Item("@PH_PY001D");
                oDS_PH_PY001E = oForm.DataSources.DBDataSources.Item("@PH_PY001E");
                oDS_PH_PY001F = oForm.DataSources.DBDataSources.Item("@PH_PY001F");
                oDS_PH_PY001G = oForm.DataSources.DBDataSources.Item("@PH_PY001G");
                oDS_PH_PY001H = oForm.DataSources.DBDataSources.Item("@PH_PY001H");
                oDS_PH_PY001J = oForm.DataSources.DBDataSources.Item("@PH_PY001J");
                oDS_PH_PY001K = oForm.DataSources.DBDataSources.Item("@PH_PY001K");
                oDS_PH_PY001L = oForm.DataSources.DBDataSources.Item("@PH_PY001L");
                oDS_PH_PY001M = oForm.DataSources.DBDataSources.Item("@PH_PY001M");
                oDS_PH_PY001N = oForm.DataSources.DBDataSources.Item("@PH_PY001N");
                oDS_PH_PY001P = oForm.DataSources.DBDataSources.Item("@PH_PY001P");
                oDS_PH_PY001Q = oForm.DataSources.DBDataSources.Item("@PH_PY001Q");
                oDS_PH_PY001R = oForm.DataSources.DBDataSources.Item("@PH_PY001R");

                oMat1 = oForm.Items.Item("Mat1").Specific; //@PH_PY001B
                oMat2 = oForm.Items.Item("Mat2").Specific; //@PH_PY001C
                oMat3 = oForm.Items.Item("Mat3").Specific; //@PH_PY001D
                oMat4 = oForm.Items.Item("Mat4").Specific; //@PH_PY001E
                oMat5 = oForm.Items.Item("Mat5").Specific; //@PH_PY001F
                oMat6 = oForm.Items.Item("Mat6").Specific; //@PH_PY001G
                oMat7 = oForm.Items.Item("Mat7").Specific; //@PH_PY001H
                oMat8 = oForm.Items.Item("Mat8").Specific; //@PH_PY001J
                oMat9 = oForm.Items.Item("Mat9").Specific; //@PH_PY001K
                oMat10 = oForm.Items.Item("Mat10").Specific; //@PH_PY001L
                oMat11 = oForm.Items.Item("Mat11").Specific; //@PH_PY001M
                oMat12 = oForm.Items.Item("Mat12").Specific; //@PH_PY001N
                oMat13 = oForm.Items.Item("Mat13").Specific; //@PH_PY001P
                oMat14 = oForm.Items.Item("Mat14").Specific; //@PH_PY001Q
                oMat15 = oForm.Items.Item("Mat15").Specific; //@PH_PY001Q

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();
                oMat2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat2.AutoResizeColumns();
                oMat3.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat3.AutoResizeColumns();
                oMat4.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat4.AutoResizeColumns();
                oMat5.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat5.AutoResizeColumns();
                oMat6.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat6.AutoResizeColumns();
                oMat7.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat7.AutoResizeColumns();
                oMat8.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat8.AutoResizeColumns();
                oMat9.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat9.AutoResizeColumns();
                oMat10.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat10.AutoResizeColumns();
                oMat11.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat11.AutoResizeColumns();
                oMat12.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat12.AutoResizeColumns();
                oMat13.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat13.AutoResizeColumns();
                oMat14.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat14.AutoResizeColumns();
                oMat15.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat15.AutoResizeColumns();

                //----------------------------------------------------------------------------------------------
                // 기본사항
                //----------------------------------------------------------------------------------------------
                oForm.AutoManaged = true;
                PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

                dataHelpClass.AutoManaged(oForm, "Code,FullName,CLTCOD");

                //사업장
                oCombo = oForm.Items.Item("CLTCOD").Specific;
                //    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
                //    Call SetReDataCombo(oForm, sQry, oCombo)
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 성
                oCombo = oForm.Items.Item("sex").Specific;
                oCombo.ValidValues.Add("M", "남자");
                oCombo.ValidValues.Add("F", "여자");
                //    oCombo.Select 0, psk_Index
                oForm.Items.Item("sex").DisplayDesc = true;

                // 음양력
                oCombo = oForm.Items.Item("CldType").Specific;
                oCombo.ValidValues.Add("S", "양력");
                oCombo.ValidValues.Add("R", "음력");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("CldType").DisplayDesc = true;

                // 최종학력
                oCombo = oForm.Items.Item("lastEduc").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P123' AND U_UseYN = 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("lastEduc").DisplayDesc = true;

                //부서
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                //담당
                oForm.Items.Item("RspCode").DisplayDesc = true;

                //반
                oForm.Items.Item("ClsCode").DisplayDesc = true;

                // 재직구분
                oCombo = oForm.Items.Item("status").Specific;
                sQry = "SELECT statusID,name FROM [OHST] ORDER BY statusID";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("status").DisplayDesc = true;

                // 직책
                oCombo = oForm.Items.Item("position").Specific;
                sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("position").DisplayDesc = true;

                // 직원구분
                oCombo = oForm.Items.Item("JIGTYP").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN = 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("JIGTYP").DisplayDesc = true;

                // 전문직호칭
                oCombo = oForm.Items.Item("CallName").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P158' AND U_UseYN = 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("CallName").DisplayDesc = true;

                // 혼인여부
                oCombo = oForm.Items.Item("martStat").Specific;
                oCombo.ValidValues.Add("N", "미혼");
                oCombo.ValidValues.Add("Y", "기혼");
                //    oCombo.Select 0, psk_Index
                oForm.Items.Item("martStat").DisplayDesc = true;

                // 국적
                oCombo = oForm.Items.Item("brthCntr").Specific;
                sQry = "SELECT Code,Name FROM [OCRY] ORDER BY Code";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("brthCntr").DisplayDesc = true;

                // 근무형태
                oCombo = oForm.Items.Item("ShiftDat").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN = 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("ShiftDat").DisplayDesc = true;

                // 근무조
                oCombo = oForm.Items.Item("GNMUJO").Specific;
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN = 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("GNMUJO").DisplayDesc = true;

                //퇴직급여 계약 구분
                oCombo = oForm.Items.Item("DBDCDIV").Specific;
                oCombo.ValidValues.Add("10", "DB형");
                oCombo.ValidValues.Add("20", "DC형");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("DBDCDIV").DisplayDesc = true;

                //비상연락처 대상 관계
                oCombo = oForm.Items.Item("EmRel").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("EmRel").DisplayDesc = true;

                //근무지사업장
                oCombo = oForm.Items.Item("BPLID2").Specific;
                sQry = "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
                oForm.Items.Item("BPLID2").DisplayDesc = true;

                oCheck = oForm.Items.Item("Inform01").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;
                oCheck = oForm.Items.Item("Inform02").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;
                oCheck = oForm.Items.Item("Inform03").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;
                oCheck = oForm.Items.Item("Inform04").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;
                oCheck = oForm.Items.Item("Inform05").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;
                oCheck = oForm.Items.Item("Inform06").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                //----------------------------------------------------------------------------------------------
                //급여기본
                //----------------------------------------------------------------------------------------------

                // 1.급여형태
                oCombo = oForm.Items.Item("PAYTYP").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");

                // 2.직급형태
                oCombo = oForm.Items.Item("JIGCOD").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN = 'Y' ORDER BY U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");

                // 3.원가구분
                oCombo = oForm.Items.Item("ACCCOD").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P140' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");

                // 4.내국인구분
                oCombo = oForm.Items.Item("INTGBN").Specific;
                oDS_PH_PY001A.SetValue("U_INTGBN", 0, "1");
                oCombo.ValidValues.Add("1", "내국인");
                oCombo.ValidValues.Add("9", "외국인");

                // 5.거주자구분
                oCombo = oForm.Items.Item("DWEGBN").Specific;
                oDS_PH_PY001A.SetValue("U_DWEGBN", 0, "1");
                oCombo.ValidValues.Add("1", "거주자");
                oCombo.ValidValues.Add("2", "비거주자");

                // 6.외국인단일세율적용
                oCombo = oForm.Items.Item("FRGTAX").Specific;
                oDS_PH_PY001A.SetValue("U_FRGTAX", 0, "2");
                oCombo.ValidValues.Add("1", "적용함");
                oCombo.ValidValues.Add("2", "적용안함");

                // 7.급여지급대상
                oCombo = oForm.Items.Item("PAYSEL").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P213' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");

                // 8.건강보험해외파견경감율
                oCombo = oForm.Items.Item("MEDFRG").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P193' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");

                // 9.세대주여부
                oCombo = oForm.Items.Item("HUSMAN").Specific;
                oCombo.ValidValues.Add("1", "세대주");
                oCombo.ValidValues.Add("2", "세대원");

                //    If oForm.Mode <> fm_FIND_MODE And oCombo.ValidValues.Count > 1 Then
                //        oCombo.Select 1, psk_Index
                //    End If

                // 고용보험여부
                oCheck = oForm.Items.Item("GBHSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 상여지급여부
                oCheck = oForm.Items.Item("BNSSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";

                // 세액계산여부
                oCheck = oForm.Items.Item("TAXSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 국외비과세여부
                oCheck = oForm.Items.Item("FRGSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 생산비과세여부
                oCheck = oForm.Items.Item("BX1SEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 연말정산신고
                oCheck = oForm.Items.Item("JSNSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 배우자유무
                oCheck = oForm.Items.Item("BAEWOO").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 부녀자공제대상
                oCheck = oForm.Items.Item("MZBURI").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 고용보험 납부대상자
                oCheck = oForm.Items.Item("NJCGBN").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 본인장애자유무
                oCheck = oForm.Items.Item("BJNGAE").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                // 국민연금연장신청여부
                oCheck = oForm.Items.Item("GBHSEL").Specific;
                oCheck.ValOn = "Y";
                oCheck.ValOff = "N";
                oCheck.Checked = false;

                //은행1
                oCombo = oForm.Items.Item("BANK1").Specific;
                sQry = "SELECT BankCode, BankName From ODSC";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("BANK1").DisplayDesc = true;

                //은행2
                oCombo = oForm.Items.Item("BANK2").Specific;
                sQry = "SELECT BankCode, BankName From ODSC";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("BANK2").DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //급여기본(Mat1)
                //----------------------------------------------------------------------------------------------
                oColumn = oMat1.Columns.Item("FILD01");
                oColumn.Editable = true;
                oColumn = oMat1.Columns.Item("FILD02");
                oColumn.Editable = true;
                //----------------------------------------------------------------------------------------------
                //급여기본(Mat2)
                //----------------------------------------------------------------------------------------------
                oColumn = oMat2.Columns.Item("FILD01");
                oColumn.Editable = true;
                oColumn = oMat2.Columns.Item("FILD02");
                oColumn.Editable = true;
                //----------------------------------------------------------------------------------------------
                //가족사항(Mat3)
                //----------------------------------------------------------------------------------------------

                // 생일구분
                oColumn = oMat3.Columns.Item("BIRGBN");
                oColumn.ValidValues.Add("S", "양력");
                oColumn.ValidValues.Add("R", "음력");
                oColumn.DisplayDesc = true;

                // 학력
                oColumn = oMat3.Columns.Item("FamSch");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P123' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //// 본인과의 관계
                oColumn = oMat3.Columns.Item("FamGun");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                // 동거여부
                oColumn = oMat3.Columns.Item("JsnType");
                oColumn.ValidValues.Add("N", "N");
                oColumn.ValidValues.Add("Y", "Y");
                oColumn.DisplayDesc = true;

                // 관계코드
                oColumn = oMat3.Columns.Item("ChkCod");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P159' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                // 내외국인
                oColumn = oMat3.Columns.Item("ChkInt");
                oColumn.ValidValues.Add("1", "내국인");
                oColumn.ValidValues.Add("9", "외국인");
                oColumn.DisplayDesc = true;

                //----------------------------------------------------------------------------------------------
                //경력사항(Mat4)
                //----------------------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------------------
                //자격사항(Mat5)
                //----------------------------------------------------------------------------------------------
                //자격/면허
                oColumn = oMat5.Columns.Item("License");

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P125' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //----------------------------------------------------------------------------------------------
                //발령사항(Mat6)
                //----------------------------------------------------------------------------------------------
                //발령구분
                oColumn = oMat6.Columns.Item("appType");

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P160' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //사업장
                oColumn = oMat6.Columns.Item("CLTCOD");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //부서
                oColumn = oMat6.Columns.Item("TeamCode");
                sQry = " SELECT U_Code, U_CodeNm, U_Comment2 FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim() + "(" + oRecordSet.Fields.Item(2).Value.ToString().Trim() + ")");
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //담당
                oColumn = oMat6.Columns.Item("RspCode");
                sQry = " SELECT U_Code, U_CodeNm ,U_Comment2 FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim() + "(" + oRecordSet.Fields.Item(2).Value.ToString().Trim() + ")");
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //직원구분
                oColumn = oMat6.Columns.Item("JIGTYP");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN = 'Y' ORDER BY CAST(U_Code AS NUMERIC) ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //직책
                oColumn = oMat6.Columns.Item("position");

                sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                // 호칭
                oColumn = oMat6.Columns.Item("CallName");

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P158' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //직급
                oColumn = oMat6.Columns.Item("JIGCOD");

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //----------------------------------------------------------------------------------------------
                //동호회(Mat7)
                //----------------------------------------------------------------------------------------------
                //동호회
                oColumn = oMat7.Columns.Item("Club");

                sQry = "  SELECT      U_Code,";
                sQry += "             U_CodeNm";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code='P156'";
                sQry += "             AND U_Char1 = '" + dataHelpClass.User_BPLID() + "'";
                sQry += "             AND U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Seq";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //// 직책
                //    Set oColumn = oMat7.Columns("Position")
                //    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
                //    oRecordSet.DoQuery sQry
                //    If oRecordSet.RecordCount > 0 Then
                //        For i = 0 To oRecordSet.RecordCount - 1
                //            oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
                //            oRecordSet.MoveNext
                //        Next i
                //    End If
                //    oColumn.DisplayDesc = True

                //----------------------------------------------------------------------------------------------
                //학력사항(Mat8)
                //----------------------------------------------------------------------------------------------
                //학위
                oColumn = oMat8.Columns.Item("degree");

                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P147' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //이수구분
                oColumn = oMat8.Columns.Item("complete");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P148' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //교육사항(Mat9)
                //----------------------------------------------------------------------------------------------
                //분류
                oColumn = oMat9.Columns.Item("type");
                sQry = "SELECT edType, Name FROM OHED";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //----------------------------------------------------------------------------------------------
                //상벌사항(Mat10)
                //----------------------------------------------------------------------------------------------
                //구분
                oColumn = oMat10.Columns.Item("Category");
                oColumn.ValidValues.Add("1", "상");
                oColumn.ValidValues.Add("2", "벌");
                oColumn.DisplayDesc = true;

                //상벌종류
                oColumn = oMat10.Columns.Item("Type");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P130' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;

                //----------------------------------------------------------------------------------------------
                //여권사항(Mat11)
                //----------------------------------------------------------------------------------------------
                //구분
                oColumn = oMat11.Columns.Item("TypeCode");
                oColumn.ValidValues.Add("1", "여권");
                oColumn.ValidValues.Add("2", "비자");

                //여권종류
                oColumn = oMat11.Columns.Item("PassType");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P151' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //노조이력(Mat12)
                //----------------------------------------------------------------------------------------------
                //직책(노조)
                oColumn = oMat12.Columns.Item("Position");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P152' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //경조사항(Mat13)
                //----------------------------------------------------------------------------------------------
                //관계
                oColumn = oMat13.Columns.Item("Relation");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //사고현황(Mat14)
                //----------------------------------------------------------------------------------------------
                //사고구분
                oColumn = oMat14.Columns.Item("type");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P153' AND U_UseYN = 'Y' ORDER BY U_Code ";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oColumn.DisplayDesc = true;
                //----------------------------------------------------------------------------------------------
                //기타사항
                //----------------------------------------------------------------------------------------------

                // 주택유무
                oCombo = oForm.Items.Item("House").Specific;
                oCombo.ValidValues.Add("N", "No");
                oCombo.ValidValues.Add("Y", "Yes");
                oForm.Items.Item("House").DisplayDesc = true;

                // 주거구분
                oCombo = oForm.Items.Item("HType").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P139' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("HType").DisplayDesc = true;

                //생활정도
                optBtn = oForm.Items.Item("ClassJ").Specific;
                optBtn.GroupWith("ClassS");

                optBtn = oForm.Items.Item("ClassH").Specific;
                optBtn.GroupWith("ClassS");

                //군별
                oCombo = oForm.Items.Item("military").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P138' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("military").DisplayDesc = true;

                //병과
                oCombo = oForm.Items.Item("Arm").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P149' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("Arm").DisplayDesc = true;

                //계급
                oCombo = oForm.Items.Item("Rank").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P150' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("Rank").DisplayDesc = true;

                //역종
                oCombo = oForm.Items.Item("StatusCd").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P136' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("StatusCd").DisplayDesc = true;

                //전역구분
                oCombo = oForm.Items.Item("DisType").Specific;
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P137' AND U_UseYN = 'Y' ORDER BY U_Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("DisType").DisplayDesc = true;

                //사내출입증
                oCombo = oForm.Items.Item("military").Specific;

                oCombo.ValidValues.Add("1", "전지역");
                oCombo.ValidValues.Add("2", "제한지역");
                oCombo.ValidValues.Add("3", "출입불가");

                oForm.Items.Item("military").DisplayDesc = true;

                //종류 - 외국어 1
                oCombo = oForm.Items.Item("foreLan1").Specific;
                sQry = " SELECT Code, Name FROM [OLNG]  ORDER BY Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("foreLan1").DisplayDesc = true;

                //종류 - 외국어 2
                oCombo = oForm.Items.Item("foreLan2").Specific;
                sQry = " SELECT Code, Name FROM [OLNG]  ORDER BY Code";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                        oRecordSet.MoveNext();
                    }
                }
                oForm.Items.Item("foreLan2").DisplayDesc = true;

                //혈액형1
                oCombo = oForm.Items.Item("BloodTy1").Specific;
                oCombo.ValidValues.Add("1", "RH+");
                oCombo.ValidValues.Add("2", "RH-");
                oForm.Items.Item("BloodTy1").DisplayDesc = true;

                //혈액형2
                oCombo = oForm.Items.Item("BloodTy2").Specific;
                oCombo.ValidValues.Add("1", "A");
                oCombo.ValidValues.Add("2", "B");
                oCombo.ValidValues.Add("3", "O");
                oCombo.ValidValues.Add("4", "AB");
                oForm.Items.Item("BloodTy2").DisplayDesc = true;

                //색맹여부
                oCombo = oForm.Items.Item("colBlind").Specific;
                oCombo.ValidValues.Add("N", "No");
                oCombo.ValidValues.Add("Y", "Yes");
                oForm.Items.Item("colBlind").DisplayDesc = true;

                //검진구분
                oCombo = oForm.Items.Item("MediExam").Specific;
                oCombo.ValidValues.Add("1", "일반검진");
                oCombo.ValidValues.Add("9", "특수검진");

                oForm.Items.Item("MediExam").DisplayDesc = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY001_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(optBtn);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY001_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY001_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY001_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY001_FormItemEnabled();
                    PH_PY001_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY001_FormItemEnabled();

                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY001_FormItemEnabled()
        {   
            string sQry;
            int i;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //oForm.Items.Item("Code").Enabled = True
                    oForm.Items.Item("FullName").Enabled = false;
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("Btn01").Visible = false;

                    oDS_PH_PY001A.SetValue("U_brthCntr", 0, "KR"); //국적

                    oDS_PH_PY001A.SetValue("U_CldType", 0, "S"); // 음양력

                    // 4.내국인구분
                    //oCombo = oForm.Items.Item("INTGBN").Specific;
                    if (oForm.Items.Item("INTGBN").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("INTGBN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    // 5.거주자구분
                    //oCombo = oForm.Items.Item("DWEGBN").Specific;
                    if (oForm.Items.Item("DWEGBN").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("DWEGBN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    // 6.외국인단일세율적용
                    //oCombo = oForm.Items.Item("FRGTAX").Specific;
                    if (oForm.Items.Item("FRGTAX").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("FRGTAX").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    oDS_PH_PY001A.SetValue("U_GBHSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_BNSSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_TAXSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_JSNSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_Inform01", 0, "N");
                    oDS_PH_PY001A.SetValue("U_Inform02", 0, "N");
                    oDS_PH_PY001A.SetValue("U_Inform03", 0, "N");
                    oDS_PH_PY001A.SetValue("U_Inform04", 0, "N");
                    oDS_PH_PY001A.SetValue("U_Inform05", 0, "N");
                    oDS_PH_PY001A.SetValue("U_Inform06", 0, "N");


                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //기본사항 - 부서 (사업장에 따른 부서변경)
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)))
                    {
                        sQry = "  SELECT      '' AS [Code], ";
                        sQry += "             '' AS [Name],";
                        sQry += "             -1 AS [Seq]";
                        sQry += " UNION ALL";
                        sQry += " SELECT      U_Code AS [Code], ";
                        sQry += "             U_CodeNm AS [Name],";
                        sQry += "             U_Seq AS [Seq]";
                        sQry += " FROM        [@PS_HR200L] ";
                        sQry += " WHERE       Code = '1' ";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    Seq";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");
                    }

                    //담당 (사업장에 따른 담당변경)
                    if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim()))
                    {
                        sQry = "  SELECT      '' AS [Code], ";
                        sQry += "             '' AS [Name],";
                        sQry += "             -1 AS [Seq]";
                        sQry += " UNION ALL";
                        sQry += " SELECT      U_Code AS [Code], ";
                        sQry += "             U_CodeNm AS [Name],";
                        sQry += "             U_Seq AS [Seq]";
                        sQry += " FROM        [@PS_HR200L] ";
                        sQry += " WHERE       Code = '2' ";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    Seq";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");
                    }

                    //반 (사업장에 따른 반변경)
                    if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("ClsCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value))
                    {
                        sQry = "  SELECT      '' AS [Code], ";
                        sQry += "             '' AS [Name],";
                        sQry += "             -1 AS [Seq]";
                        sQry += " UNION ALL";
                        sQry += " SELECT      U_Code AS [Code], ";
                        sQry += "             U_CodeNm AS [Name],";
                        sQry += "             U_Seq AS [Seq]";
                        sQry += " FROM        [@PS_HR200L] ";
                        sQry += " WHERE       Code = '9' ";
                        sQry += "             AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    Seq";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "N");
                    }
                    
                    oForm.ActiveItem = "Code"; //커서를 첫번째 ITEM으로 지정
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //사진 이미지 초기화(2017.05.25 송명규)
                    oForm.Items.Item("Pic1").Specific.Picture = "";

                    //oForm.Items.Item("Code").Enabled = True
                    oForm.Items.Item("FullName").Enabled = true;
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("Btn01").Visible = true;

                    // 4.내국인구분
                    if (oForm.Items.Item("INTGBN").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("INTGBN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    // 5.거주자구분
                    if (oForm.Items.Item("DWEGBN").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("DWEGBN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    // 6.외국인단일세율적용
                    if (oForm.Items.Item("FRGTAX").Specific.ValidValues.Count > 0)
                    {
                        oForm.Items.Item("FRGTAX").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    oDS_PH_PY001A.SetValue("U_GBHSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_BNSSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_TAXSEL", 0, "Y");
                    oDS_PH_PY001A.SetValue("U_JSNSEL", 0, "Y");

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //근무조
                    if (oForm.Items.Item("GNMUJO").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("GNMUJO").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("GNMUJO").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("GNMUJO").Specific.ValidValues.Add("", "-");
                        oForm.Items.Item("GNMUJO").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("FullName").Enabled = false;
                    oForm.Items.Item("CLTCOD").Enabled = false;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        oForm.EnableMenu("1281", false); //문서찾기
                    }
                    else
                    {
                        oForm.EnableMenu("1281", true); //문서찾기
                    }

                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY001_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Matirx 행 추가
        /// </summary>
        private void PH_PY001_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                //[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001B.GetValue("U_FILD01", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY001B.InsertRecord(oRow);
                        }
                        oDS_PH_PY001B.Offset = oRow;
                        oDS_PH_PY001B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001B.SetValue("U_FILD01", oRow, "");
                        oDS_PH_PY001B.SetValue("U_FILD02", oRow, "");
                        oDS_PH_PY001B.SetValue("U_FILD03", oRow, "0");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001B.Offset = oRow - 1;
                        oDS_PH_PY001B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001B.SetValue("U_FILD01", oRow - 1, "");
                        oDS_PH_PY001B.SetValue("U_FILD02", oRow - 1, "");
                        oDS_PH_PY001B.SetValue("U_FILD03", oRow - 1, "0");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY001B.Offset = oRow;
                    oDS_PH_PY001B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001B.SetValue("U_FILD01", oRow, "");
                    oDS_PH_PY001B.SetValue("U_FILD02", oRow, "");
                    oDS_PH_PY001B.SetValue("U_FILD03", oRow, "0");
                    oMat1.LoadFromDataSource();
                }

                //[Mat2]
                oMat2.FlushToDataSource();
                oRow = oMat2.VisualRowCount;

                if (oMat2.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001C.GetValue("U_FILD01", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001C.Size <= oMat2.VisualRowCount)
                        {
                            oDS_PH_PY001C.InsertRecord(oRow);
                        }
                        oDS_PH_PY001C.Offset = oRow;
                        oDS_PH_PY001C.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001C.SetValue("U_FILD01", oRow, "");
                        oDS_PH_PY001C.SetValue("U_FILD02", oRow, "");
                        oDS_PH_PY001C.SetValue("U_FILD03", oRow, "0");
                        oMat2.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001C.Offset = oRow - 1;
                        oDS_PH_PY001C.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001C.SetValue("U_FILD01", oRow - 1, "");
                        oDS_PH_PY001C.SetValue("U_FILD02", oRow - 1, "");
                        oDS_PH_PY001C.SetValue("U_FILD03", oRow - 1, "0");
                        oMat2.LoadFromDataSource();
                    }
                }
                else if (oMat2.VisualRowCount == 0)
                {
                    oDS_PH_PY001C.Offset = oRow;
                    oDS_PH_PY001C.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001C.SetValue("U_FILD01", oRow, "");
                    oDS_PH_PY001C.SetValue("U_FILD02", oRow, "");
                    oDS_PH_PY001C.SetValue("U_FILD03", oRow, "0");
                    oMat2.LoadFromDataSource();
                }

                //[Mat3]
                oMat3.FlushToDataSource();
                oRow = oMat3.VisualRowCount;

                if (oMat3.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001D.GetValue("U_FamNam", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001D.Size <= oMat3.VisualRowCount)
                        {
                            oDS_PH_PY001D.InsertRecord(oRow);
                        }
                        oDS_PH_PY001D.Offset = oRow;
                        oDS_PH_PY001D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001D.SetValue("U_FamNam", oRow, "");
                        oDS_PH_PY001D.SetValue("U_FamPer", oRow, "");
                        oDS_PH_PY001D.SetValue("U_BIRGBN", oRow, "");
                        oDS_PH_PY001D.SetValue("U_BIRDAT", oRow, "");
                        oDS_PH_PY001D.SetValue("U_FamSch", oRow, "");
                        oDS_PH_PY001D.SetValue("U_FamJob", oRow, "");
                        oDS_PH_PY001D.SetValue("U_FamGun", oRow, "");
                        oDS_PH_PY001D.SetValue("U_JsnType", oRow, "");
                        oDS_PH_PY001D.SetValue("U_ChkCod", oRow, "");
                        oDS_PH_PY001D.SetValue("U_ChkInt", oRow, "");
                        oDS_PH_PY001D.SetValue("U_CHKTAX", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_ChkBas", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_ChkJan", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_ChkBoH", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_ChkChl", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_ChkEdu", oRow, "N");
                        oDS_PH_PY001D.SetValue("U_Remark", oRow, "");
                        oMat3.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001D.Offset = oRow - 1;
                        oDS_PH_PY001D.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001D.SetValue("U_FamNam", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_FamPer", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_BIRGBN", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_BIRDAT", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_FamSch", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_FamJob", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_FamGun", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_JsnType", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_ChkCod", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_ChkInt", oRow - 1, "");
                        oDS_PH_PY001D.SetValue("U_CHKTAX", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_ChkBas", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_ChkJan", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_ChkBoH", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_ChkChl", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_ChkEdu", oRow - 1, "N");
                        oDS_PH_PY001D.SetValue("U_Remark", oRow - 1, "");
                        oMat3.LoadFromDataSource();
                    }
                }
                else if (oMat3.VisualRowCount == 0)
                {
                    oDS_PH_PY001D.Offset = oRow;
                    oDS_PH_PY001D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001D.SetValue("U_FamNam", oRow, "");
                    oDS_PH_PY001D.SetValue("U_FamPer", oRow, "");
                    oDS_PH_PY001D.SetValue("U_BIRGBN", oRow, "");
                    oDS_PH_PY001D.SetValue("U_BIRDAT", oRow, "");
                    oDS_PH_PY001D.SetValue("U_FamSch", oRow, "");
                    oDS_PH_PY001D.SetValue("U_FamJob", oRow, "");
                    oDS_PH_PY001D.SetValue("U_FamGun", oRow, "");
                    oDS_PH_PY001D.SetValue("U_JsnType", oRow, "");
                    oDS_PH_PY001D.SetValue("U_ChkCod", oRow, "");
                    oDS_PH_PY001D.SetValue("U_ChkInt", oRow, "");
                    oDS_PH_PY001D.SetValue("U_CHKTAX", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_ChkBas", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_ChkJan", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_ChkBoH", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_ChkChl", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_ChkEdu", oRow, "N");
                    oDS_PH_PY001D.SetValue("U_Remark", oRow, "");
                    oMat3.LoadFromDataSource();
                }

                //[Mat4]
                oMat4.FlushToDataSource();
                oRow = oMat4.VisualRowCount;

                if (oMat4.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001E.GetValue("U_employer", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001E.Size <= oMat4.VisualRowCount)
                        {
                            oDS_PH_PY001E.InsertRecord(oRow);
                        }
                        oDS_PH_PY001E.Offset = oRow;
                        oDS_PH_PY001E.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001E.SetValue("U_fromDate", oRow, "");
                        oDS_PH_PY001E.SetValue("U_toDate", oRow, "");
                        oDS_PH_PY001E.SetValue("U_employer", oRow, "");
                        oDS_PH_PY001E.SetValue("U_depart", oRow, "");
                        oDS_PH_PY001E.SetValue("U_position", oRow, "");
                        oDS_PH_PY001E.SetValue("U_remarks", oRow, "");
                        oMat4.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001E.Offset = oRow - 1;
                        oDS_PH_PY001E.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001E.SetValue("U_fromDate", oRow - 1, "");
                        oDS_PH_PY001E.SetValue("U_toDate", oRow - 1, "");
                        oDS_PH_PY001E.SetValue("U_employer", oRow - 1, "");
                        oDS_PH_PY001E.SetValue("U_depart", oRow - 1, "");
                        oDS_PH_PY001E.SetValue("U_position", oRow - 1, "");
                        oDS_PH_PY001E.SetValue("U_remarks", oRow - 1, "");
                        oMat4.LoadFromDataSource();
                    }
                }
                else if (oMat4.VisualRowCount == 0)
                {
                    oDS_PH_PY001E.Offset = oRow;
                    oDS_PH_PY001E.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001E.SetValue("U_fromDate", oRow, "");
                    oDS_PH_PY001E.SetValue("U_toDate", oRow, "");
                    oDS_PH_PY001E.SetValue("U_employer", oRow, "");
                    oDS_PH_PY001E.SetValue("U_depart", oRow, "");
                    oDS_PH_PY001E.SetValue("U_position", oRow, "");
                    oDS_PH_PY001E.SetValue("U_remarks", oRow, "");
                    oMat4.LoadFromDataSource();
                }

                //[Mat5]
                oMat5.FlushToDataSource();
                oRow = oMat5.VisualRowCount;

                if (oMat5.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001F.GetValue("U_License", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001F.Size <= oMat5.VisualRowCount)
                        {
                            oDS_PH_PY001F.InsertRecord(oRow);
                        }
                        oDS_PH_PY001F.Offset = oRow;
                        oDS_PH_PY001F.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001F.SetValue("U_License", oRow, "");
                        oDS_PH_PY001F.SetValue("U_Grade", oRow, "");
                        oDS_PH_PY001F.SetValue("U_LCNumber", oRow, "");
                        oDS_PH_PY001F.SetValue("U_IssuInst", oRow, "");
                        oDS_PH_PY001F.SetValue("U_acquDate", oRow, "");
                        oDS_PH_PY001F.SetValue("U_loseDate", oRow, "");
                        oDS_PH_PY001F.SetValue("U_Require", oRow, "");
                        oDS_PH_PY001F.SetValue("U_LCPay", oRow, "");
                        oDS_PH_PY001F.SetValue("U_PayDate", oRow, "");
                        oMat5.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001F.Offset = oRow - 1;
                        oDS_PH_PY001F.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001F.SetValue("U_License", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_Grade", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_LCNumber", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_IssuInst", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_acquDate", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_loseDate", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_Require", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_LCPay", oRow - 1, "");
                        oDS_PH_PY001F.SetValue("U_PayDate", oRow - 1, "");
                        oMat5.LoadFromDataSource();
                    }
                }
                else if (oMat5.VisualRowCount == 0)
                {
                    oDS_PH_PY001F.Offset = oRow;
                    oDS_PH_PY001F.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001F.SetValue("U_License", oRow, "");
                    oDS_PH_PY001F.SetValue("U_Grade", oRow, "");
                    oDS_PH_PY001F.SetValue("U_LCNumber", oRow, "");
                    oDS_PH_PY001F.SetValue("U_IssuInst", oRow, "");
                    oDS_PH_PY001F.SetValue("U_acquDate", oRow, "");
                    oDS_PH_PY001F.SetValue("U_loseDate", oRow, "");
                    oDS_PH_PY001F.SetValue("U_Require", oRow, "");
                    oDS_PH_PY001F.SetValue("U_LCPay", oRow, "");
                    oDS_PH_PY001F.SetValue("U_PayDate", oRow, "");
                    oMat5.LoadFromDataSource();
                }

                //[Mat6]
                oMat6.FlushToDataSource();
                oRow = oMat6.VisualRowCount;

                if (oMat6.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001G.GetValue("U_appNum", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001G.Size <= oMat6.VisualRowCount)
                        {
                            oDS_PH_PY001G.InsertRecord(oRow);
                        }
                        oDS_PH_PY001G.Offset = oRow;
                        oDS_PH_PY001G.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001G.SetValue("U_appDate", oRow, "");
                        oDS_PH_PY001G.SetValue("U_appType", oRow, "");
                        oDS_PH_PY001G.SetValue("U_appNum", oRow, "");
                        oDS_PH_PY001G.SetValue("U_TeamCode", oRow, "");
                        oDS_PH_PY001G.SetValue("U_RspCode", oRow, "");
                        oDS_PH_PY001G.SetValue("U_JIGTYP", oRow, "");
                        oDS_PH_PY001G.SetValue("U_position", oRow, "");
                        oDS_PH_PY001G.SetValue("U_CallName", oRow, "");
                        oDS_PH_PY001G.SetValue("U_JIGCOD", oRow, "");
                        oDS_PH_PY001G.SetValue("U_HOBONG", oRow, "");
                        oDS_PH_PY001G.SetValue("U_STDAMT", oRow, "0");
                        oMat6.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001G.Offset = oRow - 1;
                        oDS_PH_PY001G.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001G.SetValue("U_appDate", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_appType", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_appNum", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_TeamCode", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_RspCode", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_JIGTYP", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_position", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_CallName", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_JIGCOD", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_HOBONG", oRow - 1, "");
                        oDS_PH_PY001G.SetValue("U_STDAMT", oRow - 1, "0");
                        oMat6.LoadFromDataSource();
                    }
                }
                else if (oMat6.VisualRowCount == 0)
                {
                    oDS_PH_PY001G.Offset = oRow;
                    oDS_PH_PY001G.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001G.SetValue("U_appDate", oRow, "");
                    oDS_PH_PY001G.SetValue("U_appType", oRow, "");
                    oDS_PH_PY001G.SetValue("U_appNum", oRow, "");
                    oDS_PH_PY001G.SetValue("U_TeamCode", oRow, "");
                    oDS_PH_PY001G.SetValue("U_RspCode", oRow, "");
                    oDS_PH_PY001G.SetValue("U_JIGTYP", oRow, "");
                    oDS_PH_PY001G.SetValue("U_position", oRow, "");
                    oDS_PH_PY001G.SetValue("U_CallName", oRow, "");
                    oDS_PH_PY001G.SetValue("U_JIGCOD", oRow, "");
                    oDS_PH_PY001G.SetValue("U_HOBONG", oRow, "");
                    oDS_PH_PY001G.SetValue("U_STDAMT", oRow, "0");
                    oMat6.LoadFromDataSource();
                }

                //[Mat7]
                oMat7.FlushToDataSource();
                oRow = oMat7.VisualRowCount;

                if (oMat7.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001H.GetValue("U_InDate", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001H.Size <= oMat7.VisualRowCount)
                        {
                            oDS_PH_PY001H.InsertRecord(oRow);
                        }
                        oDS_PH_PY001H.Offset = oRow;
                        oDS_PH_PY001H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001H.SetValue("U_Club", oRow, "");
                        oDS_PH_PY001H.SetValue("U_InDate", oRow, "");
                        oDS_PH_PY001H.SetValue("U_OutDate", oRow, "");
                        oDS_PH_PY001H.SetValue("U_Position", oRow, "");
                        oDS_PH_PY001H.SetValue("U_Comments", oRow, "");
                        oMat7.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001H.Offset = oRow - 1;
                        oDS_PH_PY001H.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001H.SetValue("U_Club", oRow - 1, "");
                        oDS_PH_PY001H.SetValue("U_InDate", oRow - 1, "");
                        oDS_PH_PY001H.SetValue("U_OutDate", oRow - 1, "");
                        oDS_PH_PY001H.SetValue("U_Position", oRow - 1, "");
                        oDS_PH_PY001H.SetValue("U_Comments", oRow - 1, "");
                        oMat7.LoadFromDataSource();
                    }
                }
                else if (oMat7.VisualRowCount == 0)
                {
                    oDS_PH_PY001H.Offset = oRow;
                    oDS_PH_PY001H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001H.SetValue("U_Club", oRow, "");
                    oDS_PH_PY001H.SetValue("U_InDate", oRow, "");
                    oDS_PH_PY001H.SetValue("U_OutDate", oRow, "");
                    oDS_PH_PY001H.SetValue("U_Position", oRow, "");
                    oDS_PH_PY001H.SetValue("U_Comments", oRow, "");
                    oMat7.LoadFromDataSource();
                }

                //[Mat8]
                oMat8.FlushToDataSource();
                oRow = oMat8.VisualRowCount;

                if (oMat8.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001J.GetValue("U_School", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001J.Size <= oMat8.VisualRowCount)
                        {
                            oDS_PH_PY001J.InsertRecord(oRow);
                        }
                        oDS_PH_PY001J.Offset = oRow;
                        oDS_PH_PY001J.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001J.SetValue("U_fromDate", oRow, "");
                        oDS_PH_PY001J.SetValue("U_toDate", oRow, "");
                        oDS_PH_PY001J.SetValue("U_School", oRow, "");
                        oDS_PH_PY001J.SetValue("U_depart", oRow, "");
                        oDS_PH_PY001J.SetValue("U_degree", oRow, "");
                        oDS_PH_PY001J.SetValue("U_complete", oRow, "");
                        oMat8.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001J.Offset = oRow - 1;
                        oDS_PH_PY001J.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001J.SetValue("U_fromDate", oRow - 1, "");
                        oDS_PH_PY001J.SetValue("U_toDate", oRow - 1, "");
                        oDS_PH_PY001J.SetValue("U_School", oRow - 1, "");
                        oDS_PH_PY001J.SetValue("U_depart", oRow - 1, "");
                        oDS_PH_PY001J.SetValue("U_degree", oRow - 1, "");
                        oDS_PH_PY001J.SetValue("U_complete", oRow - 1, "");
                        oMat8.LoadFromDataSource();
                    }
                }
                else if (oMat8.VisualRowCount == 0)
                {
                    oDS_PH_PY001J.Offset = oRow;
                    oDS_PH_PY001J.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001J.SetValue("U_fromDate", oRow, "");
                    oDS_PH_PY001J.SetValue("U_toDate", oRow, "");
                    oDS_PH_PY001J.SetValue("U_School", oRow, "");
                    oDS_PH_PY001J.SetValue("U_depart", oRow, "");
                    oDS_PH_PY001J.SetValue("U_degree", oRow, "");
                    oDS_PH_PY001J.SetValue("U_complete", oRow, "");
                    oMat8.LoadFromDataSource();
                }

                //[Mat9]
                oMat9.FlushToDataSource();
                oRow = oMat9.VisualRowCount;

                if (oMat9.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001K.GetValue("U_major", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001K.Size <= oMat9.VisualRowCount)
                        {
                            oDS_PH_PY001K.InsertRecord(oRow);
                        }
                        oDS_PH_PY001K.Offset = oRow;
                        oDS_PH_PY001K.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001K.SetValue("U_type", oRow, "");
                        oDS_PH_PY001K.SetValue("U_major", oRow, "");
                        oDS_PH_PY001K.SetValue("U_fromDate", oRow, "");
                        oDS_PH_PY001K.SetValue("U_toDate", oRow, "");
                        oDS_PH_PY001K.SetValue("U_institut", oRow, "");
                        oDS_PH_PY001K.SetValue("U_diploma", oRow, "");
                        oDS_PH_PY001K.SetValue("U_EduExp", oRow, "");
                        oDS_PH_PY001K.SetValue("U_TraExp", oRow, "");
                        oMat9.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001K.Offset = oRow - 1;
                        oDS_PH_PY001K.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001K.SetValue("U_type", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_major", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_fromDate", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_toDate", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_institut", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_diploma", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_EduExp", oRow - 1, "");
                        oDS_PH_PY001K.SetValue("U_TraExp", oRow - 1, "");
                        oMat9.LoadFromDataSource();
                    }
                }
                else if (oMat9.VisualRowCount == 0)
                {
                    oDS_PH_PY001K.Offset = oRow;
                    oDS_PH_PY001K.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001K.SetValue("U_type", oRow, "");
                    oDS_PH_PY001K.SetValue("U_major", oRow, "");
                    oDS_PH_PY001K.SetValue("U_fromDate", oRow, "");
                    oDS_PH_PY001K.SetValue("U_toDate", oRow, "");
                    oDS_PH_PY001K.SetValue("U_institut", oRow, "");
                    oDS_PH_PY001K.SetValue("U_diploma", oRow, "");
                    oDS_PH_PY001K.SetValue("U_EduExp", oRow, "");
                    oDS_PH_PY001K.SetValue("U_TraExp", oRow, "");
                    oMat9.LoadFromDataSource();
                }

                //[Mat10]
                oMat10.FlushToDataSource();
                oRow = oMat10.VisualRowCount;

                if (oMat10.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001L.GetValue("U_Basis", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001L.Size <= oMat10.VisualRowCount)
                        {
                            oDS_PH_PY001L.InsertRecord(oRow);
                        }
                        oDS_PH_PY001L.Offset = oRow;
                        oDS_PH_PY001L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001L.SetValue("U_Date", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Category", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Type", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Basis", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Text", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Reward", oRow, "");
                        oDS_PH_PY001L.SetValue("U_Reduce", oRow, "");
                        oDS_PH_PY001L.SetValue("U_addition", oRow, "");
                        oDS_PH_PY001L.SetValue("U_fromYM", oRow, "");
                        oDS_PH_PY001L.SetValue("U_toYM", oRow, "");
                        oMat10.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001L.Offset = oRow - 1;
                        oDS_PH_PY001L.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001L.SetValue("U_Date", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Category", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Type", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Basis", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Text", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Reward", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_Reduce", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_addition", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_fromYM", oRow - 1, "");
                        oDS_PH_PY001L.SetValue("U_toYM", oRow - 1, "");
                        oMat10.LoadFromDataSource();
                    }
                }
                else if (oMat10.VisualRowCount == 0)
                {
                    oDS_PH_PY001L.Offset = oRow;
                    oDS_PH_PY001L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001L.SetValue("U_Date", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Category", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Type", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Basis", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Text", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Reward", oRow, "");
                    oDS_PH_PY001L.SetValue("U_Reduce", oRow, "");
                    oDS_PH_PY001L.SetValue("U_addition", oRow, "");
                    oDS_PH_PY001L.SetValue("U_fromYM", oRow, "");
                    oDS_PH_PY001L.SetValue("U_toYM", oRow, "");
                    oMat10.LoadFromDataSource();
                }

                //[Mat11]
                oMat11.FlushToDataSource();
                oRow = oMat11.VisualRowCount;

                if (oMat11.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001M.GetValue("U_Passport", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001M.Size <= oMat11.VisualRowCount)
                        {
                            oDS_PH_PY001M.InsertRecord(oRow);
                        }
                        oDS_PH_PY001M.Offset = oRow;
                        oDS_PH_PY001M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001M.SetValue("U_Passport", oRow, "");
                        oDS_PH_PY001M.SetValue("U_IssueDat", oRow, "");
                        oDS_PH_PY001M.SetValue("U_ExpirDat", oRow, "");
                        oDS_PH_PY001M.SetValue("U_PassType", oRow, "");
                        oDS_PH_PY001M.SetValue("U_Comments", oRow, "");
                        oMat11.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001M.Offset = oRow - 1;
                        oDS_PH_PY001M.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001M.SetValue("U_Passport", oRow - 1, "");
                        oDS_PH_PY001M.SetValue("U_IssueDat", oRow - 1, "");
                        oDS_PH_PY001M.SetValue("U_ExpirDat", oRow - 1, "");
                        oDS_PH_PY001M.SetValue("U_PassType", oRow - 1, "");
                        oDS_PH_PY001M.SetValue("U_Comments", oRow - 1, "");
                        oMat11.LoadFromDataSource();
                    }
                }
                else if (oMat11.VisualRowCount == 0)
                {
                    oDS_PH_PY001M.Offset = oRow;
                    oDS_PH_PY001M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001M.SetValue("U_Passport", oRow, "");
                    oDS_PH_PY001M.SetValue("U_IssueDat", oRow, "");
                    oDS_PH_PY001M.SetValue("U_ExpirDat", oRow, "");
                    oDS_PH_PY001M.SetValue("U_PassType", oRow, "");
                    oDS_PH_PY001M.SetValue("U_Comments", oRow, "");

                    oMat11.LoadFromDataSource();
                }

                //[Mat12]
                oMat12.FlushToDataSource();
                oRow = oMat12.VisualRowCount;

                if (oMat12.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001N.GetValue("U_Period", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001N.Size <= oMat12.VisualRowCount)
                        {
                            oDS_PH_PY001N.InsertRecord(oRow);
                        }
                        oDS_PH_PY001N.Offset = oRow;
                        oDS_PH_PY001N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001N.SetValue("U_fromDate", oRow, "");
                        oDS_PH_PY001N.SetValue("U_toDate", oRow, "");
                        oDS_PH_PY001N.SetValue("U_Period", oRow, "");
                        oDS_PH_PY001N.SetValue("U_Position", oRow, "");
                        oDS_PH_PY001N.SetValue("U_Comments", oRow, "");
                        oMat12.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001N.Offset = oRow - 1;
                        oDS_PH_PY001N.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001N.SetValue("U_fromDate", oRow - 1, "");
                        oDS_PH_PY001N.SetValue("U_toDate", oRow - 1, "");
                        oDS_PH_PY001N.SetValue("U_Period", oRow - 1, "");
                        oDS_PH_PY001N.SetValue("U_Position", oRow - 1, "");
                        oDS_PH_PY001N.SetValue("U_Comments", oRow - 1, "");
                        oMat12.LoadFromDataSource();
                    }
                }
                else if (oMat12.VisualRowCount == 0)
                {
                    oDS_PH_PY001N.Offset = oRow;
                    oDS_PH_PY001N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001N.SetValue("U_fromDate", oRow, "");
                    oDS_PH_PY001N.SetValue("U_toDate", oRow, "");
                    oDS_PH_PY001N.SetValue("U_Period", oRow, "");
                    oDS_PH_PY001N.SetValue("U_Position", oRow, "");
                    oDS_PH_PY001N.SetValue("U_Comments", oRow, "");
                    oMat12.LoadFromDataSource();
                }

                //[Mat13]
                oMat13.FlushToDataSource();
                oRow = oMat13.VisualRowCount;

                if (oMat13.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001P.GetValue("U_Date", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001P.Size <= oMat13.VisualRowCount)
                        {
                            oDS_PH_PY001P.InsertRecord(oRow);
                        }
                        oDS_PH_PY001P.Offset = oRow;
                        oDS_PH_PY001P.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001P.SetValue("U_Date", oRow, "");
                        oDS_PH_PY001P.SetValue("U_Relation", oRow, "");
                        oDS_PH_PY001P.SetValue("U_Person", oRow, "");
                        oDS_PH_PY001P.SetValue("U_ComExpen", oRow, "");
                        oDS_PH_PY001P.SetValue("U_empExpen", oRow, "");
                        oDS_PH_PY001P.SetValue("U_Comments", oRow, "");
                        oMat13.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001P.Offset = oRow - 1;
                        oDS_PH_PY001P.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001P.SetValue("U_Date", oRow - 1, "");
                        oDS_PH_PY001P.SetValue("U_Relation", oRow - 1, "");
                        oDS_PH_PY001P.SetValue("U_Person", oRow - 1, "");
                        oDS_PH_PY001P.SetValue("U_ComExpen", oRow - 1, "");
                        oDS_PH_PY001P.SetValue("U_empExpen", oRow - 1, "");
                        oDS_PH_PY001P.SetValue("U_Comments", oRow - 1, "");
                        oMat13.LoadFromDataSource();
                    }
                }
                else if (oMat13.VisualRowCount == 0)
                {
                    oDS_PH_PY001P.Offset = oRow;
                    oDS_PH_PY001P.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001P.SetValue("U_Date", oRow, "");
                    oDS_PH_PY001P.SetValue("U_Relation", oRow, "");
                    oDS_PH_PY001P.SetValue("U_Person", oRow, "");
                    oDS_PH_PY001P.SetValue("U_ComExpen", oRow, "");
                    oDS_PH_PY001P.SetValue("U_empExpen", oRow, "");
                    oDS_PH_PY001P.SetValue("U_Comments", oRow, "");
                    oMat13.LoadFromDataSource();
                }

                //[Mat14]
                oMat14.FlushToDataSource();
                oRow = oMat14.VisualRowCount;

                if (oMat14.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001Q.GetValue("U_injury", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001Q.Size <= oMat14.VisualRowCount)
                        {
                            oDS_PH_PY001Q.InsertRecord(oRow);
                        }
                        oDS_PH_PY001Q.Offset = oRow;
                        oDS_PH_PY001Q.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001Q.SetValue("U_type", oRow, "");
                        oDS_PH_PY001Q.SetValue("U_fromDate", oRow, "");
                        oDS_PH_PY001Q.SetValue("U_toDate", oRow, "");
                        oDS_PH_PY001Q.SetValue("U_injury", oRow, "");
                        oDS_PH_PY001Q.SetValue("U_details", oRow, "");
                        oMat14.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001Q.Offset = oRow - 1;
                        oDS_PH_PY001Q.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY001Q.SetValue("U_type", oRow - 1, "");
                        oDS_PH_PY001Q.SetValue("U_fromDate", oRow - 1, "");
                        oDS_PH_PY001Q.SetValue("U_toDate", oRow - 1, "");
                        oDS_PH_PY001Q.SetValue("U_injury", oRow - 1, "");
                        oDS_PH_PY001Q.SetValue("U_details", oRow - 1, "");
                        oMat14.LoadFromDataSource();
                    }
                }
                else if (oMat14.VisualRowCount == 0)
                {
                    oDS_PH_PY001Q.Offset = oRow;
                    oDS_PH_PY001Q.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001Q.SetValue("U_type", oRow, "");
                    oDS_PH_PY001Q.SetValue("U_fromDate", oRow, "");
                    oDS_PH_PY001Q.SetValue("U_toDate", oRow, "");
                    oDS_PH_PY001Q.SetValue("U_injury", oRow, "");
                    oDS_PH_PY001Q.SetValue("U_details", oRow, "");
                    oMat14.LoadFromDataSource();
                }

                //[Mat15]
                oMat15.FlushToDataSource();
                oRow = oMat15.VisualRowCount;

                if (oMat15.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY001R.GetValue("U_CarNum", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY001R.Size <= oMat15.VisualRowCount)
                        {
                            oDS_PH_PY001R.InsertRecord(oRow);
                        }
                        oDS_PH_PY001R.Offset = oRow;
                        oDS_PH_PY001R.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001R.SetValue("U_CarType", oRow, "");
                        oDS_PH_PY001R.SetValue("U_CarNum", oRow, "");
                        oDS_PH_PY001R.SetValue("U_InDate", oRow, "");
                        oDS_PH_PY001R.SetValue("U_ETC12", oRow, "");
                        oMat15.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY001R.Offset = oRow - 1;
                        oDS_PH_PY001R.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY001R.SetValue("U_CarType", oRow, "");
                        oDS_PH_PY001R.SetValue("U_CarNum", oRow, "");
                        oDS_PH_PY001R.SetValue("U_InDate", oRow, "");
                        oDS_PH_PY001R.SetValue("U_ETC12", oRow, "");
                        oMat15.LoadFromDataSource();
                    }
                }
                else if (oMat15.VisualRowCount == 0)
                {
                    oDS_PH_PY001R.Offset = oRow;
                    oDS_PH_PY001R.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY001R.SetValue("U_CarType", oRow, "");
                    oDS_PH_PY001R.SetValue("U_CarNum", oRow, "");
                    oDS_PH_PY001R.SetValue("U_InDate", oRow, "");
                    oDS_PH_PY001R.SetValue("U_ETC12", oRow, "");
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY001_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY001'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY001_DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY001_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            string sQry;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //----------------------------------------------------------------------------------
                //기본사항 탭
                //----------------------------------------------------------------------------------
                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("Code", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사원번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //사원번호 중복 체크 (PH_PY001A)
                    sQry = "SELECT Code FROM [@PH_PY001A] WHERE Code = '" + oDS_PH_PY001A.GetValue("Code", 0).Trim() + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("[PH_PY001A] 이미 존재하는 사원번호입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return functionReturnValue;
                    }

                    //사원번호 중복 체크(OHEM)
                    sQry = "SELECT U_MSTCOD FROM OHEM WHERE U_MSTCOD = '" + oDS_PH_PY001A.GetValue("Code", 0).Trim() + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("[OHEM] 이미 존재하는 사원번호입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return functionReturnValue;
                    }

                    // Code = Name 동일
                    oDS_PH_PY001A.SetValue("Name", 0, oDS_PH_PY001A.GetValue("Code", 0));
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_CLTCOD", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_StartDat", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("입사일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("startDat").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_lNameKO", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("한글성은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("lNameKO").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_fNameKO", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("한글명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("fNameKO").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //주민번호 체크
                if (!string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_govID", 0).Trim()))
                {
                    if (dataHelpClass.GovIDCheck(oDS_PH_PY001A.GetValue("U_govID", 0).Trim()) == false)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("잘못된 주민번호입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        oForm.Items.Item("govID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return functionReturnValue;
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("주민번호 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("govID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_birthDat", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("생년월일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("birthDat").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_fNameKO", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("한글명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("fNameKO").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                
                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_sex", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("성별은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("sex").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_status", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("재직구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 1)
                    {
                        oForm.PaneLevel = 1;
                    }
                    oForm.Items.Item("status").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_position", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("직책은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 1)
                    {
                        oForm.PaneLevel = 1;
                    }
                    oForm.Items.Item("position").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_brthCntr", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("국적은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 1)
                    {
                        oForm.PaneLevel = 1;
                    }

                    oForm.Items.Item("brthCntr").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //----------------------------------------------------------------------------------
                //급여기본 탭
                //----------------------------------------------------------------------------------
                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_PAYSEL", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("급여지급대상은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 2)
                    {
                        oForm.Items.Item("FLD02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }

                    oForm.Items.Item("PAYSEL").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_STDAMT", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("급여기본금은 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 2)
                    {
                        oForm.PaneLevel = 2;
                    }

                    oForm.Items.Item("STDAMT").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_HUSMAN", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("세대주여부는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    if (oForm.PaneLevel != 2)
                    {
                        oForm.PaneLevel = 2;
                    }

                    oForm.Items.Item("HUSMAN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //----------------------------------------------------------------------------------
                //급여기본 탭 - 고정수당
                //----------------------------------------------------------------------------------
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat1.Columns.Item("FILD01").Cells.Item(i).Specific.Value))
                        {
                            if (string.IsNullOrEmpty(oMat1.Columns.Item("FILD02").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 2)
                                {
                                    oForm.PaneLevel = 2;
                                }
                                PSH_Globals.SBO_Application.SetStatusBarMessage("수당명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat1.Columns.Item("FILD02").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //급여기본 탭 - 고정공제
                //----------------------------------------------------------------------------------
                if (oMat2.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat2.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat2.Columns.Item("FILD01").Cells.Item(i).Specific.Value))
                        {
                            if (string.IsNullOrEmpty(oMat2.Columns.Item("FILD02").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 2)
                                {
                                    oForm.PaneLevel = 2;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("공제명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat2.Columns.Item("FILD02").Cells.Item(2).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //기타사항 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //가족사항 탭
                //----------------------------------------------------------------------------------
                //가족 이름이 등록시...주민번호 입력 체크
                if (oMat3.VisualRowCount > 0)
                {
                    for (i = 1; i <= oMat3.VisualRowCount; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat3.Columns.Item("FamNam").Cells.Item(i).Specific.Value))
                        {
                            if (string.IsNullOrEmpty(oMat3.Columns.Item("FamPer").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 3)
                                {
                                    oForm.PaneLevel = 3;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("입력된 가족의 주민등록번호는 필수입니다., bmt_Short, True");
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //경력사항 탭
                //----------------------------------------------------------------------------------
                if (oMat4.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat4.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat4.Columns.Item("employer").Cells.Item(i).Specific.Value))
                        {
                            if (string.IsNullOrEmpty(oMat4.Columns.Item("fromDate").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 4)
                                {
                                    oForm.PaneLevel = 4;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("기간시작은 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat4.Columns.Item(i).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }

                            if (string.IsNullOrEmpty(oMat4.Columns.Item("toDate").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 4)
                                {
                                    oForm.PaneLevel = 4;
                                }
                                PSH_Globals.SBO_Application.SetStatusBarMessage("기간종료는 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat4.Columns.Item("toDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //자격사항탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //발령사항 탭
                //----------------------------------------------------------------------------------
                if (oMat6.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat6.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat6.Columns.Item("appType").Cells.Item(i).Specific.Value))
                        {

                            if (string.IsNullOrEmpty(oMat6.Columns.Item("appDate").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 6)
                                {
                                    oForm.PaneLevel = 6;
                                }
                                PSH_Globals.SBO_Application.SetStatusBarMessage("발령일자는 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat6.Columns.Item("appDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }

                            if (string.IsNullOrEmpty(oMat6.Columns.Item("appNum").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 6)
                                {
                                    oForm.PaneLevel = 6;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("발령번호는 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat6.Columns.Item(i).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //동호회 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //항력사항 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //교육사항 탭
                //----------------------------------------------------------------------------------
                if (oMat9.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat9.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat9.Columns.Item("major").Cells.Item(i).Specific.Value))
                        {
                            if (string.IsNullOrEmpty(oMat9.Columns.Item("type").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 9)
                                {
                                    oForm.PaneLevel = 9;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("교육사항 - 분류는 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat9.Columns.Item("type").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }

                            if (string.IsNullOrEmpty(oMat9.Columns.Item("fromDate").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 9)
                                {
                                    oForm.PaneLevel = 9;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("교육사항 -  기간시작은 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat9.Columns.Item("fromDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }

                            if (string.IsNullOrEmpty(oMat9.Columns.Item("toDate").Cells.Item(i).Specific.Value))
                            {
                                if (oForm.PaneLevel != 9)
                                {
                                    oForm.PaneLevel = 9;
                                }

                                PSH_Globals.SBO_Application.SetStatusBarMessage("교육사항 - 기간종료는 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat9.Columns.Item("toDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }

                //----------------------------------------------------------------------------------
                //상벌사항 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //여권사항 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //노조이력 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //경조사항 탭
                //----------------------------------------------------------------------------------

                //----------------------------------------------------------------------------------
                //사고현황 탭
                //----------------------------------------------------------------------------------

                //DI API
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //OHEM 등록
                    if (PH_PY001_OHEM_DI_ADD() == false)
                    {
                        return functionReturnValue;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    if (PH_PY001_OHEM_DI_UPDATE() == false)
                    {
                        return functionReturnValue;
                    }
                }

                oMat1.FlushToDataSource();
                oMat2.FlushToDataSource();
                oMat3.FlushToDataSource();
                oMat4.FlushToDataSource();
                oMat5.FlushToDataSource();
                oMat6.FlushToDataSource();
                oMat7.FlushToDataSource();
                oMat8.FlushToDataSource();
                oMat9.FlushToDataSource();
                oMat10.FlushToDataSource();
                oMat11.FlushToDataSource();
                oMat12.FlushToDataSource();
                oMat13.FlushToDataSource();
                oMat14.FlushToDataSource();
                oMat15.FlushToDataSource();

                //Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY001B.Size > 1)
                {
                    oDS_PH_PY001B.RemoveRecord(oDS_PH_PY001B.Size - 1);
                }
                    
            
                if (oDS_PH_PY001C.Size > 1)
                {
                    oDS_PH_PY001C.RemoveRecord(oDS_PH_PY001C.Size - 1);
                }

                if (oDS_PH_PY001D.Size > 1)
                {
                    oDS_PH_PY001D.RemoveRecord(oDS_PH_PY001D.Size - 1);
                }

                if (oDS_PH_PY001E.Size > 1)
                {
                    oDS_PH_PY001E.RemoveRecord(oDS_PH_PY001E.Size - 1);
                }

                if (oDS_PH_PY001F.Size > 1)
                {
                    oDS_PH_PY001F.RemoveRecord(oDS_PH_PY001F.Size - 1);
                }

                if (oDS_PH_PY001G.Size > 1)
                {
                    oDS_PH_PY001G.RemoveRecord(oDS_PH_PY001G.Size - 1);
                }

                if (oDS_PH_PY001H.Size > 1)
                {
                    oDS_PH_PY001H.RemoveRecord(oDS_PH_PY001H.Size - 1);
                }

                if (oDS_PH_PY001J.Size > 1)
                {
                    oDS_PH_PY001J.RemoveRecord(oDS_PH_PY001J.Size - 1);
                }

                if (oDS_PH_PY001K.Size > 1)
                {
                    oDS_PH_PY001K.RemoveRecord(oDS_PH_PY001K.Size - 1);
                }

                if (oDS_PH_PY001L.Size > 1)
                {
                    oDS_PH_PY001L.RemoveRecord(oDS_PH_PY001L.Size - 1);
                }

                if (oDS_PH_PY001M.Size > 1)
                {
                    oDS_PH_PY001M.RemoveRecord(oDS_PH_PY001M.Size - 1);
                }

                if (oDS_PH_PY001N.Size > 1)
                {
                    oDS_PH_PY001N.RemoveRecord(oDS_PH_PY001N.Size - 1);
                }

                if (oDS_PH_PY001P.Size > 1)
                {
                    oDS_PH_PY001P.RemoveRecord(oDS_PH_PY001P.Size - 1);
                }

                if (oDS_PH_PY001Q.Size > 1)
                {
                    oDS_PH_PY001Q.RemoveRecord(oDS_PH_PY001Q.Size - 1);
                }

                if (oDS_PH_PY001Q.Size > 1)
                {
                    oDS_PH_PY001R.RemoveRecord(oDS_PH_PY001R.Size - 1);
                }

                oMat1.LoadFromDataSource();
                oMat2.LoadFromDataSource();
                oMat3.LoadFromDataSource();
                oMat4.LoadFromDataSource();
                oMat5.LoadFromDataSource();
                oMat6.LoadFromDataSource();
                oMat7.LoadFromDataSource();
                oMat8.LoadFromDataSource();
                oMat9.LoadFromDataSource();
                oMat10.LoadFromDataSource();
                oMat11.LoadFromDataSource();
                oMat12.LoadFromDataSource();
                oMat13.LoadFromDataSource();
                oMat14.LoadFromDataSource();
                oMat15.LoadFromDataSource();

                functionReturnValue = true;

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY001_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY001_MTX01, 메트릭스에 데이터 로드
        /// </summary>
        private void PH_PY001_MTX01()
        {
            int i;
            string sQry;
            string errCode = string.Empty;

            string Param01;
            string Param02;
            string Param03;
            string Param04;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("Param01").Specific.Value;
                Param02 = oForm.Items.Item("Param01").Specific.Value;
                Param03 = oForm.Items.Item("Param01").Specific.Value;
                Param04 = oForm.Items.Item("Param01").Specific.Value;

                sQry = "SELECT 10";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY001B.InsertRecord(i);
                    }
                    oDS_PH_PY001B.Offset = i;
                    oDS_PH_PY001B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY001B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
                    oRecordSet.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY001_Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY001_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            string errCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.Get_ReData("Canceled", "DocEntry", "[@PH_PY001A]", "'" & oForm.Items.Item("DocEntry").Specific.Value & "'", "") == "Y")
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

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }
        
        /// <summary>
        /// 인사기본and급여 저장 DI
        /// </summary>
        /// <returns></returns>
        private bool PH_PY001_OHEM_DI_ADD()
        {
            bool functionReturnValue = false;
            int i;
            string sQry;
            string errCode = string.Empty;
            int errDiCode = 0;
            string errDiMsg = string.Empty;

            SAPbobsCOM.EmployeesInfo oOHEM = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo);
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_Globals.SBO_Application.SetStatusBarMessage(oForm.Items.Item("Code").Specific.Value + "사원마스터가 생성중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);

            try
            {
                oOHEM.UserFields.Fields.Item("U_MSTCOD").Value = oForm.Items.Item("Code").Specific.Value; //사원번호
                oOHEM.UserFields.Fields.Item("U_CLTCOD").Value = oForm.Items.Item("CLTCOD").Specific.Value; //사업장
                oOHEM.UserFields.Fields.Item("U_GRPDAT").Value = DateTime.ParseExact(oForm.Items.Item("GrpDat").Specific.Value, "yyyyMMdd", null); //그룹입사일
                oOHEM.StartDate = DateTime.ParseExact(oForm.Items.Item("startDat").Specific.Value, "yyyyMMdd", null); //입사일자

                if (!string.IsNullOrEmpty(oForm.Items.Item("RetDat").Specific.Value))
                {
                    oOHEM.UserFields.Fields.Item("U_RETDAT").Value = DateTime.ParseExact(oForm.Items.Item("RetDat").Specific.Value, "yyyyMMdd", null); //중간정산일 (값이 있을 때만 입력)
                }

                oOHEM.LastName = oForm.Items.Item("lNameKO").Specific.Value; //한글 성
                oOHEM.FirstName = oForm.Items.Item("fNameKO").Specific.Value; //한글 명
                oOHEM.UserFields.Fields.Item("U_FULLNAME").Value = oOHEM.LastName.Trim() + oOHEM.FirstName.Trim(); //성명(풀네임)
                oOHEM.UserFields.Fields.Item("U_lNameTW").Value = oForm.Items.Item("lNameTW").Specific.Value; //한자 성
                oOHEM.UserFields.Fields.Item("U_fNameTW").Value = oForm.Items.Item("fNameTW").Specific.Value; //한자 명
                oOHEM.UserFields.Fields.Item("U_lNameEN").Value = oForm.Items.Item("lNameEN").Specific.Value; //영문 성
                oOHEM.UserFields.Fields.Item("U_fNameEN").Value = oForm.Items.Item("fNameEN").Specific.Value; //영문 명
                oOHEM.IdNumber = oForm.Items.Item("govID").Specific.Value; //주민번호(govID)

                if (oForm.Items.Item("sex").Specific.Value == "M") //성별
                {
                    oOHEM.Gender = SAPbobsCOM.BoGenderTypes.gt_Male; //"F"
                }
                else
                {
                    oOHEM.Gender = SAPbobsCOM.BoGenderTypes.gt_Female;
                }

                oOHEM.DateOfBirth = DateTime.ParseExact(oForm.Items.Item("birthDat").Specific.Value, "yyyyMMdd", null); //생년월일
                oOHEM.UserFields.Fields.Item("U_lastEduc").Value = oForm.Items.Item("lastEduc").Specific.Value; //최종학력
                oOHEM.eMail = oForm.Items.Item("eMail").Specific.Value; //전자메일
                oOHEM.HomeStreet = oForm.Items.Item("Address").Specific.Value; //주소
                oOHEM.HomeBlock = oForm.Items.Item("OriginAd").Specific.Value; //본적
                oOHEM.Remarks = oForm.Items.Item("MainJob").Specific.Value; //주요업무
                oOHEM.UserFields.Fields.Item("U_TeamCode").Value = oForm.Items.Item("TeamCode").Specific.Value; //부서
                oOHEM.UserFields.Fields.Item("U_RspCode").Value = oForm.Items.Item("RspCode").Specific.Value; //담당
                oOHEM.UserFields.Fields.Item("U_ClsCode").Value = oForm.Items.Item("ClsCode").Specific.Value; //반
                oOHEM.HomePhone = oForm.Items.Item("homeTel").Specific.Value; //전화번호
                oOHEM.OfficePhone = oForm.Items.Item("OffcTel").Specific.Value; //사무실전화
                oOHEM.MobilePhone = oForm.Items.Item("mobile").Specific.Value; //휴대전화
                oOHEM.StatusCode = Convert.ToInt16(oForm.Items.Item("status").Specific.Value); //재직구분

                if (Convert.ToInt16(oForm.Items.Item("status").Specific.Value) == 5) //재직구분이 - 퇴직일 경우
                {
                    oOHEM.TerminationDate = DateTime.ParseExact(oForm.Items.Item("termDate").Specific.Value, "yyyyMMdd", null); //퇴직일자
                }

                oOHEM.Position = Convert.ToInt16(oForm.Items.Item("position").Specific.Value); //직책
                oOHEM.UserFields.Fields.Item("U_JIGCOD").Value = oForm.Items.Item("JIGCOD").Specific.Value; //직급
                oOHEM.UserFields.Fields.Item("U_JIGTYP").Value = oForm.Items.Item("JIGTYP").Specific.Value; //직원구분
                oOHEM.UserFields.Fields.Item("U_CallName").Value = oForm.Items.Item("CallName").Specific.Value; //전문직호칭
                oOHEM.UserFields.Fields.Item("U_preCode").Value = oForm.Items.Item("preCode").Specific.Value; //변경전사번

                if (oForm.Items.Item("martStat").Specific.Value == "Y") //혼인구분
                {
                    oOHEM.MartialStatus = SAPbobsCOM.BoMeritalStatuses.mts_Married; // "N"
                }
                else
                {
                    oOHEM.MartialStatus = SAPbobsCOM.BoMeritalStatuses.mts_Single;
                }

                if (!string.IsNullOrEmpty(oForm.Items.Item("weddingD").Specific.Value.ToString().Trim()))
                {
                    oOHEM.UserFields.Fields.Item("U_weddingD").Value = DateTime.ParseExact(oForm.Items.Item("weddingD").Specific.Value, "yyyyMMdd", null); //결혼기념일
                }

                oOHEM.CountryOfBirth = oForm.Items.Item("brthCntr").Specific.Value; //국적
                oOHEM.UserFields.Fields.Item("U_ShiftTyp").Value = oForm.Items.Item("ShiftDat").Specific.Value; //근무형태
                oOHEM.UserFields.Fields.Item("U_GNMUJO").Value = oForm.Items.Item("GNMUJO").Specific.Value; //근무조
                oOHEM.UserFields.Fields.Item("U_KUKJUN").Value = oForm.Items.Item("KUKJUN").Specific.Value; //퇴직근전환금

                if (oMat4.VisualRowCount > 1) //경력사항 OHEM4
                {
                    for (i = 1; i <= oMat4.VisualRowCount - 1; i++)
                    {
                        oOHEM.PreviousEmpoymentInfo.Add();
                        oOHEM.PreviousEmpoymentInfo.FromDtae = DateTime.ParseExact(oMat4.Columns.Item("fromDate").Cells.Item(i).Specific.Value, "yyyyMMdd", null); //기간From
                        oOHEM.PreviousEmpoymentInfo.ToDate = DateTime.ParseExact(oMat4.Columns.Item("toDate").Cells.Item(i).Specific.Value, "yyyyMMdd", null); //기간 To
                        oOHEM.PreviousEmpoymentInfo.Employer = oMat4.Columns.Item("employer").Cells.Item(i).Specific.Value; //근무처
                        oOHEM.PreviousEmpoymentInfo.UserFields.Fields.Item("U_depart").Value = oMat4.Columns.Item("depart").Cells.Item(i).Specific.Value; //부서
                        oOHEM.PreviousEmpoymentInfo.Position = oMat4.Columns.Item("position").Cells.Item(i).Specific.Value; //직책
                        oOHEM.PreviousEmpoymentInfo.Remarks = oMat4.Columns.Item("remarks").Cells.Item(i).Specific.Value; //직장소재지
                    }
                }

                if (oMat9.VisualRowCount > 1) //교육사항 OHEM2
                {
                    for (i = 1; i <= oMat9.VisualRowCount - 1; i++)
                    {
                        sQry = "SELECT Name FROM OHED WHERE edTYPE = '" + oMat9.Columns.Item("type").Cells.Item(i).Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.RecordCount > 0)
                        {
                            oOHEM.EducationInfo.Add();
                            oOHEM.EducationInfo.EducationType = oMat9.Columns.Item("type").Cells.Item(i).Specific.Value; //분류
                            oOHEM.EducationInfo.Major = oMat9.Columns.Item("major").Cells.Item(i).Specific.Value; //교육명
                            oOHEM.EducationInfo.FromDate = DateTime.ParseExact(oMat9.Columns.Item("fromDate").Cells.Item(i).Specific.Value, "yyyyMMdd", null); //기간from
                            oOHEM.EducationInfo.ToDate = DateTime.ParseExact(oMat9.Columns.Item("toDate").Cells.Item(i).Specific.Value, "yyyyMMdd", null); //기간to
                            oOHEM.EducationInfo.Institute = oMat9.Columns.Item("institut").Cells.Item(i).Specific.Value; //교육기간
                            oOHEM.EducationInfo.Diploma = oMat9.Columns.Item("diploma").Cells.Item(i).Specific.Value; //교육장소
                            oOHEM.EducationInfo.UserFields.Fields.Item("U_EduExp").Value = oMat9.Columns.Item("EduExp").Cells.Item(i).Specific.Value; //교육비
                            oOHEM.EducationInfo.UserFields.Fields.Item("U_TraExp").Value = oMat9.Columns.Item("TraExp").Cells.Item(i).Specific.Value; //출장비
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("등록된 교육 분류가 없어, 교육사항은 저장되지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                }

                if (0 != oOHEM.Add())
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage(oForm.Items.Item("Code").Specific.Value + "사원마스터가 생성되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    oDS_PH_PY001A.SetValue("U_FullName", 0, oOHEM.UserFields.Fields.Item("U_FULLNAME").Value);
                    functionReturnValue = true;
                }
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_OHEM_DI_ADD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oOHEM);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY001_OHEM_DI_UPDATE
        /// </summary>
        /// <returns></returns>
        private bool PH_PY001_OHEM_DI_UPDATE()
        {
            bool functionReturnValue = false;
            string sQry;
            int errDiCode = 0;
            string errDiMsg = string.Empty;
            string errCode = string.Empty;
            string EmpID;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.EmployeesInfo oOHEM = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo);

            PSH_Globals.SBO_Application.SetStatusBarMessage(oForm.Items.Item("Code").Specific.Value + "사원마스터가 갱신중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);

            try
            {
                sQry = "SELECT U_Empid FROM [@PH_PY001A] WHERE Code = '" + oForm.Items.Item("Code").Specific.Value.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                EmpID = oRecordSet.Fields.Item(0).Value;

                if (string.IsNullOrEmpty(EmpID))
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oOHEM.GetByKey(Convert.ToInt32(EmpID)) == true)
                {
                    oOHEM.GetByKey(Convert.ToInt32(EmpID));
                    oOHEM.UserFields.Fields.Item("U_MSTCOD").Value = oForm.Items.Item("Code").Specific.Value; // 사원번호
                    oOHEM.UserFields.Fields.Item("U_CLTCOD").Value = oForm.Items.Item("CLTCOD").Specific.Value; //사업장
                    oOHEM.UserFields.Fields.Item("U_GRPDAT").Value = DateTime.ParseExact(oForm.Items.Item("GrpDat").Specific.Value, "yyyyMMdd", null); //그룹입사일
                    oOHEM.StartDate = DateTime.ParseExact(oForm.Items.Item("startDat").Specific.Value, "yyyyMMdd", null); //입사일자

                    if (!string.IsNullOrEmpty(oForm.Items.Item("RetDat").Specific.Value))
                    {
                        oOHEM.UserFields.Fields.Item("U_RETDAT").Value = DateTime.ParseExact(oForm.Items.Item("RetDat").Specific.Value, "yyyyMMdd", null); //중간정산일
                    }
                    oOHEM.LastName = oForm.Items.Item("lNameKO").Specific.Value; //한글 성
                    oOHEM.FirstName = oForm.Items.Item("fNameKO").Specific.Value; //한글 명
                    oOHEM.UserFields.Fields.Item("U_FULLNAME").Value = oOHEM.LastName.Trim() + oOHEM.FirstName.Trim(); //성명(풀네임)
                    oOHEM.UserFields.Fields.Item("U_lNameTW").Value = oForm.Items.Item("lNameTW").Specific.Value; //한자 성
                    oOHEM.UserFields.Fields.Item("U_fNameTW").Value = oForm.Items.Item("fNameTW").Specific.Value; //한자 명
                    oOHEM.UserFields.Fields.Item("U_lNameEN").Value = oForm.Items.Item("lNameEN").Specific.Value; //영문 성
                    oOHEM.UserFields.Fields.Item("U_fNameEN").Value = oForm.Items.Item("fNameEN").Specific.Value; //영문 명
                    oOHEM.IdNumber = oForm.Items.Item("govID").Specific.Value; //주민번호(govID)

                    if (oForm.Items.Item("sex").Specific.Value == "M") //성별
                    {
                        oOHEM.Gender = SAPbobsCOM.BoGenderTypes.gt_Male;
                    }
                    else
                    {
                        oOHEM.Gender = SAPbobsCOM.BoGenderTypes.gt_Female;
                    }

                    oOHEM.DateOfBirth = DateTime.ParseExact(oForm.Items.Item("birthDat").Specific.Value, "yyyyMMdd", null); //생년월일
                    oOHEM.UserFields.Fields.Item("U_lastEduc").Value = oForm.Items.Item("lastEduc").Specific.Value; //최종학력
                    oOHEM.eMail = oForm.Items.Item("eMail").Specific.Value; //전자메일
                    oOHEM.HomeStreet = oForm.Items.Item("Address").Specific.Value; //주소
                    oOHEM.HomeBlock = oForm.Items.Item("OriginAd").Specific.Value; //본적
                    oOHEM.Remarks = oForm.Items.Item("MainJob").Specific.Value; //주요업무
                    oOHEM.UserFields.Fields.Item("U_TeamCode").Value = oForm.Items.Item("TeamCode").Specific.Value; //부서
                    oOHEM.UserFields.Fields.Item("U_RspCode").Value = oForm.Items.Item("RspCode").Specific.Value; //담당
                    oOHEM.UserFields.Fields.Item("U_ClsCode").Value = oForm.Items.Item("ClsCode").Specific.Value; //반
                    oOHEM.HomePhone = oForm.Items.Item("homeTel").Specific.Value; //전화번호
                    oOHEM.MobilePhone = oForm.Items.Item("mobile").Specific.Value; //휴대전화
                    oOHEM.StatusCode = Convert.ToInt16(oForm.Items.Item("status").Specific.Value); //재직구분

                    if (Convert.ToInt16(oForm.Items.Item("status").Specific.Value) == 5) //재직구분이 - 퇴직일 경우
                    {
                        oOHEM.TerminationDate = DateTime.ParseExact(oForm.Items.Item("termDate").Specific.Value, "yyyyMMdd", null); //퇴직일자
                    }

                    oOHEM.Position = Convert.ToInt16(oForm.Items.Item("position").Specific.Value); //직책
                    oOHEM.UserFields.Fields.Item("U_JIGCOD").Value = oForm.Items.Item("JIGCOD").Specific.Value; //직급
                    oOHEM.UserFields.Fields.Item("U_JIGTYP").Value = oForm.Items.Item("JIGTYP").Specific.Value; //직원구분
                    oOHEM.UserFields.Fields.Item("U_CallName").Value = oForm.Items.Item("CallName").Specific.Value; //전문직호칭
                    oOHEM.UserFields.Fields.Item("U_preCode").Value = oForm.Items.Item("preCode").Specific.Value; //변경전사번

                    if (!string.IsNullOrEmpty(oForm.Items.Item("weddingD").Specific.Value.ToString().Trim()))
                    {
                        oOHEM.UserFields.Fields.Item("U_weddingD").Value = DateTime.ParseExact(oForm.Items.Item("weddingD").Specific.Value, "yyyyMMdd", null); //결혼기념일
                    }

                    if (oForm.Items.Item("martStat").Specific.Value == "Y") //혼인구분
                    {
                        oOHEM.MartialStatus = SAPbobsCOM.BoMeritalStatuses.mts_Married; // "N"
                    }
                    else
                    {
                        oOHEM.MartialStatus = SAPbobsCOM.BoMeritalStatuses.mts_Single;
                    }

                    oOHEM.CountryOfBirth = oForm.Items.Item("brthCntr").Specific.Value; //국적
                    oOHEM.UserFields.Fields.Item("U_ShiftTyp").Value = oForm.Items.Item("ShiftDat").Specific.Value; //근무형태
                    oOHEM.UserFields.Fields.Item("U_GNMUJO").Value = oForm.Items.Item("GNMUJO").Specific.Value; //근무조
                    oOHEM.UserFields.Fields.Item("U_KUKJUN").Value = oForm.Items.Item("KUKJUN").Specific.Value; //퇴직금전환금

                    if (0 != oOHEM.Update())
                    {
                        PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                        errCode = "2";
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage(oForm.Items.Item("Code").Specific.Value + "사원마스터가 갱신되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        functionReturnValue = true;
                    }
                }
            }
            catch(Exception ex)
            {
                if(errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사원마스터의 사원번호(EmpID) 로드에 실패하였습니다. 전산담당에게 문의하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_OHEM_DI_UPDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oOHEM);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY001_EmpIDUpdate(DI등록 후 사번 업데이트)
        /// </summary>
        private void PH_PY001_EmpIDUpdate()
        {
            int EmpID;
            string sQry;
            string MSTCOD;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //OHEM - empID
                sQry = "SELECT Autokey FROM ONNM WHERE objectcode = '171'";

                oRecordSet.DoQuery(sQry);

                EmpID = oRecordSet.Fields.Item(0).Value - 1;

                //OHEM - empID 의 사원번호
                sQry = "SELECT U_MSTCOD FROM [OHEM] WHERE empID = '" + EmpID + "'";

                oRecordSet.DoQuery(sQry);

                MSTCOD = oRecordSet.Fields.Item(0).Value;

                //@PH_PY001A 에 empid 업데이트
                sQry = "  UPDATE  [@PH_PY001A]";
                sQry += " SET     U_empID = '" + EmpID + "'";
                sQry += " WHERE   Code = '" + MSTCOD + "'";

                oRecordSet.DoQuery(sQry);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_EmpIDUpdate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY001_Display_Hobong
        /// </summary>
        private void PH_PY001_Display_Hobong()
        {
            string sQry;
            int iRow;
            int jRow;
            
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset sRecordset = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "  SELECT  U_CSUCOD, U_CSUNAM";
                sQry += " FROM    [@PH_PY102B] T1";
                if (!string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_HOBYMM", 0).ToString().Trim()))
                {
                    sQry += " WHERE   CODE = (SELECT TOP 1 CODE FROM [@PH_PY102A] WHERE CODE <= '" + oDS_PH_PY001A.GetValue("U_HOBYMM", 0).ToString().Trim() + "' ORDER BY CODE DESC)";
                }
                else
                {
                    sQry += " WHERE   CODE = (SELECT TOP 1 CODE FROM [@PH_PY102A] ORDER BY CODE DESC)";
                }
                sQry += " AND     U_HOBUSE = 'Y'";
                sQry += " ORDER   BY U_INSLIN";
                sRecordset.DoQuery(sQry);

                if (sRecordset.RecordCount > 0)
                {

                    jRow = 0;
                    sQry = "  SELECT      TOP 1";
                    sQry += "             T0.U_STDAMT,";
                    sQry += "             T0.U_BNSAMT,";
                    sQry += "             T0.U_EXTAMT01,";
                    sQry += "             T0.U_EXTAMT02,";
                    sQry += "             T0.U_EXTAMT03,";
                    sQry += "             T0.U_EXTAMT04,";
                    sQry += "             T0.U_EXTAMT05,";
                    sQry += "             T0.U_EXTAMT06,";
                    sQry += "             T0.U_EXTAMT07,";
                    sQry += "             T0.U_EXTAMT08,";
                    sQry += "             T0.U_EXTAMT09,";
                    sQry += "             T0.U_EXTAMT10,";
                    sQry += "             T1.U_YM";
                    sQry += " FROM        [@PH_PY105B] T0";
                    sQry += "             INNER JOIN";
                    sQry += "             [@PH_PY105A] T1";
                    sQry += "                 ON T0.Code = T1.Code";
                    sQry += " WHERE       T0.U_JIGCOD = '" + oDS_PH_PY001A.GetValue("U_JIGCOD", 0).ToString().Trim() + "'";
                    sQry += "             AND T0.U_HOBCOD = '" + oDS_PH_PY001A.GetValue("U_HOBONG", 0).ToString().Trim() + "'";
                    if (!string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_HOBYMM", 0).ToString().Trim()))
                    {
                        sQry += "         AND T1.U_YM = '" + oDS_PH_PY001A.GetValue("U_HOBYMM", 0).ToString().Trim() + "'";
                    }
                    sQry += " ORDER BY    T1.U_YM DESC";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        oDS_PH_PY001A.SetValue("U_STDAMT", 0, oRecordSet.Fields.Item(0).Value);
                        oDS_PH_PY001A.SetValue("U_BNSAMT", 0, oRecordSet.Fields.Item(1).Value);
                        oDS_PH_PY001A.SetValue("U_HOBYMM", 0, oRecordSet.Fields.Item("U_YM").Value);
                        
                        for (iRow = 0; iRow <= oDS_PH_PY001B.Size - 1; iRow++)
                        {
                            if (oDS_PH_PY001B.GetValue("U_FILD01", iRow).ToString().Trim() == sRecordset.Fields.Item("U_CSUCOD").Value)
                            {
                                jRow += 1;

                                oDS_PH_PY001B.SetValue("U_FILD03", iRow, oRecordSet.Fields.Item(1 + jRow).Value);
                                sRecordset.MoveNext();
                            }
                        }
                    }
                    oForm.Items.Item("STDAMT").Update();
                    oForm.Items.Item("BNSAMT").Update();
                    oMat1.LoadFromDataSource();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_Display_Hobong_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sRecordset);
            }
        }

        /// <summary>
        /// PH_PY001_Matrix_Init, 메트릭스 초기화
        /// </summary>
        private void PH_PY001_Matrix_Init()
        {
            string sQry;
            int i;
            int cnt;
            int old_Line = 0;
            int errNum = 0;
            bool DataChk;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //수당Matrix 초기화
                cnt = oDS_PH_PY001B.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY001B.RemoveRecord(oDS_PH_PY001B.Size - 1);
                    }
                    if (cnt == 1)
                    {
                        oDS_PH_PY001B.Clear();
                    }
                }
                oMat1.LoadFromDataSource();

                //공제Matrix 초기화
                cnt = oDS_PH_PY001C.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY001C.RemoveRecord(oDS_PH_PY001C.Size - 1);
                    }
                    if (cnt == 1)
                    {
                        oDS_PH_PY001C.Clear();
                    }
                }
                oMat2.LoadFromDataSource();

                //기존 재직자의 수당항목과 비교
                DataChk = true;
                sQry = "  SELECT      T0.U_FILD01,";
                sQry += "             T0.U_FILD02";
                sQry += " FROM        [@PH_PY001A] T1";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY001B] T0";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T1.Code = (";
                sQry += "                         SELECT      TOP 1";
                sQry += "                                     Code";
                sQry += "                         FROM        [@PH_PY001A]";
                sQry += "                         WHERE       U_Status ='1'";
                sQry += "                         ORDER BY    Code";
                sQry += "                       )";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    old_Line = 0;
                    while (!oRecordSet.EoF)
                    {
                        if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value))
                        {
                            old_Line += 1;
                        }
                        oRecordSet.MoveNext();
                    }
                    DataChk = false;
                }
                //수당항목코드 조회
                sQry = "  SELECT      T0.U_CSUCOD,";
                sQry += "             T0.U_CSUNAM";
                sQry += " FROM        [@PH_PY102B] T0";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY102A] T1";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T0.U_FIXGBN = '1'"; //고정급
                sQry += "             AND Left(T0.U_CSUCOD,1) <> 'A'"; //기본급제외
                sQry += "             AND ISNULL(T0.U_INSLIN, '') <> ''"; //인사마스터 순서
                sQry += " ORDER BY    T0.Code,";
                sQry += "             CAST(T0.U_INSLIN AS INT),";
                sQry += "             CAST(T0.U_LINSEQ AS INT)";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    if (DataChk == false && old_Line < oRecordSet.RecordCount)
                    {
                        errNum = 1;
                        throw new Exception();
                    }
                    i = 1;
                    while (!oRecordSet.EoF)
                    {
                        if (i > oDS_PH_PY001B.Size)
                        {
                            oDS_PH_PY001B.InsertRecord(i - 1);
                        }
                        oDS_PH_PY001B.Offset = i - 1;
                        oDS_PH_PY001B.SetValue("U_LineNum", i - 1, Convert.ToString(i));
                        oDS_PH_PY001B.SetValue("U_FILD01", i - 1, oRecordSet.Fields.Item(0).Value); //코드
                        oDS_PH_PY001B.SetValue("U_FILD02", i - 1, oRecordSet.Fields.Item(1).Value); //명
                        oDS_PH_PY001B.SetValue("U_FILD03", i - 1, ""); //금액
                        i += 1;
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.LoadFromDataSource();

                //공제항목
                DataChk = true;
                sQry = "  SELECT      T0.U_FILD01,";
                sQry += "             T0.U_FILD02";
                sQry += " FROM        [@PH_PY001A] T1";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY001C] T0";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T1.Code = (";
                sQry += "                         SELECT      TOP 1";
                sQry += "                                      Code";
                sQry += "                         FROM        [@PH_PY001A]";
                sQry += "                         WHERE       U_Status ='1'";
                sQry += "                         ORDER BY    Code";
                sQry += "                       )";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    old_Line = 0;
                    while (!oRecordSet.EoF)
                    {
                        if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value))
                        {
                            old_Line += 1;
                        }
                        oRecordSet.MoveNext();
                    }
                    DataChk = false;
                }

                //공제항목코드읽어오기
                sQry = "  SELECT      T0.U_CSUCOD,";
                sQry += "             T0.U_CSUNAM ";
                sQry += " FROM        [@PH_PY103B] T0";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY103A] T1";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T0.U_FIXGBN = '1'"; //고정공제
                sQry += "             AND ISNULL(T0.U_INSLIN, '') <> ''"; //인사마스터순서
                sQry += "             AND Left(T0.U_CSUCOD,1)  <> 'A'"; //기본급제외
                sQry += " ORDER BY    T0.Code,";
                sQry += "             CAST(T0.U_INSLIN AS INT)";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    if (DataChk == false && old_Line < oRecordSet.RecordCount)
                    {
                        errNum = 2;
                        throw new Exception();
                    }

                    i = 1;
                    while (!oRecordSet.EoF)
                    {
                        if (i > oDS_PH_PY001C.Size)
                        {
                            oDS_PH_PY001C.InsertRecord(i - 1);
                        }
                        oDS_PH_PY001C.Offset = i - 1;
                        oDS_PH_PY001C.SetValue("U_LineNum", i - 1, Convert.ToString(i));
                        oDS_PH_PY001C.SetValue("U_FILD01", i - 1, oRecordSet.Fields.Item(0).Value); //코드
                        oDS_PH_PY001C.SetValue("U_FILD02", i - 1, oRecordSet.Fields.Item(1).Value); //명
                        oDS_PH_PY001C.SetValue("U_FILD03", i - 1, ""); //금액

                        i += 1;
                        oRecordSet.MoveNext();
                    }
                }
                oMat2.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_Matrix_Init_Error : 기존자료와 수당항목관리의 고정수당이 다릅니다. 고정수당/공제 생성으로 재직자의 고정수당을 변경 후 사용하십시오." + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_Matrix_Init_Error : 기존자료와 공제항목관리의 고정공제이 다릅니다. 고정수당/공제 생성으로 재직자의 고정공제를 변경 후 사용하십시오." + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_Matrix_Init_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// OpenFileSelectDialog 호출(쓰레드를 이용하여 비동기화)
        /// OLE 호출을 수행하려면 현재 스레드를 STA(단일 스레드 아파트) 모드로 설정해야 합니다.
        /// </summary>
        [STAThread]
        private string OpenFileSelectDialog()
        {
            string returnFileName = string.Empty;

            var thread = new System.Threading.Thread(() =>
            {
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "bmp Files|*.bmp|All Files|*.*";
                openFileDialog.FilterIndex = 1; //FilterIndex는 1부터 시작
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    returnFileName = openFileDialog.FileName;
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            return returnFileName;
        }

        /// <summary>
        /// 사진파일 업로드
        /// </summary>
        private void PH_PY001_UpLoadPic()
        {
            short errNum = 0;
            string fullFileName;
            string fileName;
            string sourcePath;
            string targetPath;
            string sourceFile;
            string targetFile;
            string newFileName; //사진파일을 사번으로 저장하기 위한 새로운 파일명용 변수

            try
            {   
                if (string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("Code", 0).ToString().Trim())) //사번입력 확인, 입력하지 않으면 다음단계 진행 불가, 파일명을 사번명으로 변경하기 위함
                {
                    errNum = 1;
                    throw new Exception();
                }

                fullFileName = OpenFileSelectDialog(); //OpenFileDialog를 쓰레드로 실행

                if (string.IsNullOrEmpty(fullFileName)) //파일을 선택하지 않고, 확인이나 취소버튼을 클릭했을 경우
                {
                    return;
                }

                fileName = System.IO.Path.GetFileName(fullFileName); //파일명
                sourcePath = System.IO.Path.GetDirectoryName(fullFileName); //파일명을 제외한 전체 경로
                targetPath = "\\\\" + PSH_Globals.SP_ODBC_IP + "\\HR_Pic"; //191.1.1.220\HR_Pic"; //목적지(서버) 경로(하드코딩은 지향하자. 수정필요)

                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                targetFile = System.IO.Path.Combine(targetPath, fileName);

                if (!System.IO.Directory.Exists(targetPath)) //폴더가 존재하지 않으면 생성
                {
                    System.IO.Directory.CreateDirectory(targetPath);
                }

                newFileName = oDS_PH_PY001A.GetValue("Code", 0).ToString().Trim() + ".BMP"; //사진파일을 사번으로 저장하기 위한 새로운 파일명
                if (fileName != newFileName) //파일이름이 사번과 다르면 사번으로 변경
                {
                    targetFile = System.IO.Path.Combine(targetPath, newFileName);
                }

                //서버에 기존파일이 존재하는지 체크
                if (System.IO.File.Exists(targetFile))
                {
                    if (PSH_Globals.SBO_Application.MessageBox("현재 사원의 사진 파일이 있습니다, 교체하시겠습니까?", 2, "Yes", "No") == 1)
                    {
                        System.IO.File.Delete(targetFile); //삭제
                    }
                    else
                    {
                        return;
                    }
                }

                //파일 복사(이미 존재하는 경우 덮어씀)
                System.IO.File.Copy(sourceFile, targetFile, true);

                PSH_Globals.SBO_Application.StatusBar.SetText("사진이 서버에 업로드 되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oDS_PH_PY001A.SetValue("U_Pic1", 0, targetFile);

            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사번 입력후 사진을 업로드 하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_UpLoadPic_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY001_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            // 접속자 ID저장
                            oDS_PH_PY001A.SetValue("U_UserSign2", 0, PSH_Globals.oCompany.UserSignature.ToString());
                            oDS_PH_PY001A.SetValue("U_UpdtProg", 0, "PH_PY001_Y");

                            if (PH_PY001_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        if (pVal.ItemUID == "Btn01")
                        {
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn02")
                    {
                        oForm.Items.Item("HOBONG").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        BubbleEvent = false;
                    }
                    if (pVal.ItemUID == "Pic2")
                    {
                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                            PH_PY001_UpLoadPic();
                        }
                    }
                    if (pVal.ItemUID == "Pic3")
                    {
                        oDS_PH_PY001A.SetValue("U_Pic1", 0, "");
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "FLD01":
                        case "FLD02":
                        case "FLD03":
                        case "FLD04":
                        case "FLD05":
                        case "FLD06":
                        case "FLD07":
                        case "FLD08":
                        case "FLD09":
                        case "FLD10":
                        case "FLD11":
                        case "FLD12":
                        case "FLD13":
                        case "FLD14":
                        case "FLD15":
                        case "FLD16":
                            oForm.PaneLevel = Convert.ToInt32(pVal.ItemUID.Substring(3, 2));
                            break;
                        case "1":
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY001_EmpIDUpdate();
                                    PH_PY001_FormItemEnabled();
                                    PH_PY001_AddMatrixRow();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY001_FormItemEnabled();
                                    PH_PY001_AddMatrixRow();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY001_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                    PH_PY001_FormItemEnabled();
                                }
                            }
                            break;
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
                        case "Mat2":
                        case "Mat3":
                        case "Mat4":
                        case "Mat5":
                        case "Mat6":
                        case "Mat7":
                        case "Mat8":
                        case "Mat9":
                        case "Mat10":
                        case "Mat11":
                        case "Mat12":
                        case "Mat13":
                        case "Mat14":
                        case "Mat15":
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
            string sQry;
            int loopCount;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        //사업장(헤더)
                        if (pVal.ItemUID == "CLTCOD")
                        {
                            //기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (loopCount = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                                oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                                oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Seq]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '1' ";
                            sQry += "             AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");
                            
                            oForm.Items.Item("TeamCode").DisplayDesc = true;
                        }
                        else if (pVal.ItemUID == "TeamCode")
                        {
                            //담당 (팀에 따른 담당변경)
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (loopCount = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                                oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                                oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Seq]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '2' ";
                            sQry += "             AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");

                            oForm.Items.Item("RspCode").DisplayDesc = true;
                        }
                        else if (pVal.ItemUID == "RspCode")
                        {
                            //반 (사업장에 따른 반변경)
                            if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                            {
                                for (loopCount = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                                {
                                    oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                                oForm.Items.Item("ClsCode").Specific.ValidValues.Add("", "");
                                oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Name]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '9' ";
                            sQry += "             AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_Char2 = '" + oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "N");

                            oForm.Items.Item("GrpDat").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        //재직구분에 따른 퇴직일자 활성화
                        if (pVal.ItemUID == "status")
                        {
                            switch (oDS_PH_PY001A.GetValue("U_status", 0))
                            {
                                case "1":
                                case "2":
                                case "3":
                                case "4":
                                    oDS_PH_PY001A.SetValue("U_termDate", 0, "");
                                    oForm.Items.Item("termDate").Enabled = false;
                                    break;
                                case "5":
                                    oForm.Items.Item("termDate").Enabled = true;
                                    break;
                            }
                        }

                        //혼인구분에 따른 결혼기념일 활성화
                        if (pVal.ItemUID == "martStat")
                        {
                            switch (oDS_PH_PY001A.GetValue("U_martStat", 0))
                            {
                                case "N":
                                    oDS_PH_PY001A.SetValue("U_weddingD", 0, "");
                                    oForm.Items.Item("weddingD").Enabled = false;
                                    break;

                                case "Y":
                                    oForm.Items.Item("weddingD").Enabled = true;
                                    break;
                            }
                        }

                        //근무형태에 따른 근무조 값
                        if (pVal.ItemUID == "ShiftDat")
                        {
                            //oCombo = oForm.Items.Item("GNMUJO").Specific;

                            if (oForm.Items.Item("GNMUJO").Specific.ValidValues.Count > 0)
                            {
                                for (loopCount = oForm.Items.Item("GNMUJO").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                                {
                                    oForm.Items.Item("GNMUJO").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P155' AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.Value.ToString().Trim() + "'";
                            sQry += " ORDER BY U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "N");

                            oForm.Items.Item("GNMUJO").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("GNMUJO").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "PAYTYP")
                        {
                            switch (oDS_PH_PY001A.GetValue("U_PAYTYP", 0).ToString().Trim())
                            {
                                case "1":
                                    oDS_PH_PY001A.SetValue("U_HOBONG", 0, "000");
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = true;
                                    break;
                                case "2":
                                    oDS_PH_PY001A.SetValue("U_HOBONG", 0, "");
                                    oForm.Items.Item("Btn02").Visible = true;
                                    oForm.Items.Item("HOBONG").Enabled = true;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                                case "3":
                                    oDS_PH_PY001A.SetValue("U_HOBONG", 0, "000");
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                                case "4":
                                    oDS_PH_PY001A.SetValue("U_HOBONG", 0, "000");
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                            }
                            oForm.Items.Item("HOBONG").Update();
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
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
                        case "Mat2":
                        case "Mat3":
                        case "Mat4":
                        case "Mat5":
                        case "Mat6":
                        case "Mat7":
                        case "Mat8":
                        case "Mat9":
                        case "Mat10":
                        case "Mat11":
                        case "Mat12":
                        case "Mat13":
                        case "Mat14":
                        case "Mat15":
                            if (pVal.Row > 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Mat1":
                                        oMat1.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat2":
                                        oMat2.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat3":
                                        oMat3.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat4":
                                        oMat4.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat5":
                                        oMat5.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat6":
                                        oMat6.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat7":
                                        oMat7.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat8":
                                        oMat8.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat9":
                                        oMat9.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat10":
                                        oMat10.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat11":
                                        oMat11.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat12":
                                        oMat12.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat13":
                                        oMat13.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat14":
                                        oMat14.SelectRow(pVal.Row, true, false);
                                        break;
                                    case "Mat15":
                                        oMat15.SelectRow(pVal.Row, true, false);
                                        break;
                                }
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
                        case "Mat2":
                        case "Mat3":
                        case "Mat4":
                        case "Mat5":
                        case "Mat6":
                        case "Mat7":
                        case "Mat8":
                        case "Mat9":
                        case "Mat10":
                        case "Mat11":
                        case "Mat12":
                        case "Mat13":
                        case "Mat14":
                        case "Mat15":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string tSex = string.Empty;
            string errCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat3" && pVal.ColUID == "FamPer")
                    {
                        if (!string.IsNullOrEmpty(oMat3.Columns.Item("FamPer").Cells.Item(pVal.Row).Specific.Value))
                        {
                            if (dataHelpClass.GovIDCheck(oMat3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) == false)
                            {
                                errCode = "1";
                                throw new Exception();
                            }
                        }
                    }

                    if (pVal.ItemUID == "govID")
                    {
                        if (!string.IsNullOrEmpty(oDS_PH_PY001A.GetValue("U_govID", 0)))
                        {
                            if (dataHelpClass.GovIDCheck(oDS_PH_PY001A.GetValue("U_govID", 0)) == true)
                            {
                                //성별구분
                                if (Convert.ToDouble(oDS_PH_PY001A.GetValue("U_govID", 0).ToString().Substring(6, 1)) == 1 || Convert.ToDouble(oDS_PH_PY001A.GetValue("U_govID", 0).ToString().Substring(6, 1)) == 3)
                                {
                                    tSex = "M";
                                }
                                else if (Convert.ToDouble(oDS_PH_PY001A.GetValue("U_govID", 0).ToString().Substring(6, 1)) == 2 || Convert.ToDouble(oDS_PH_PY001A.GetValue("U_govID", 0).ToString().Substring(6, 1)) == 4)
                                {
                                    tSex = "F";
                                }

                                oDS_PH_PY001A.SetValue("U_Sex", 0, tSex);
                            }
                            else
                            {
                                errCode = "2";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "FullName":
                                if (!string.IsNullOrEmpty(oForm.Items.Item("FullName").Specific.Value))
                                {
                                    oDS_PH_PY001A.SetValue("U_MZBURI", 0, (oDS_PH_PY001A.GetValue("U_sex", 0).ToString().Trim() == "F" ? "1" : "0"));
                                }
                                break;
                            case "HOBONG": //호봉코드
                                if (string.IsNullOrEmpty(oForm.Items.Item("HOBONG").Specific.Value))
                                {
                                    oDS_PH_PY001A.SetValue("U_YUNAMT", 0, "0");
                                    oDS_PH_PY001A.SetValue("U_STDAMT", 0, "0");
                                    oDS_PH_PY001A.SetValue("U_BNSAMT", 0, "0");
                                    PH_PY001_Matrix_Init();
                                }
                                else
                                {
                                    PH_PY001_Display_Hobong();
                                }
                                break;
                            case "YUNAMT": //연봉
                                if (Convert.ToDouble(oDS_PH_PY001A.GetValue("U_YUNAMT", 0).ToString().Trim()) > 0 && oDS_PH_PY001A.GetValue("U_PAYTYP", 0).ToString().Trim() == "1")
                                {
                                    oDS_PH_PY001A.SetValue("U_STDAMT", 0, Convert.ToString(Convert.ToDouble(oDS_PH_PY001A.GetValue("U_YUNAMT", 0).ToString().Trim()) / 20));
                                    oDS_PH_PY001A.SetValue("U_BNSAMT", 0, Convert.ToString(Convert.ToDouble(oDS_PH_PY001A.GetValue("U_YUNAMT", 0).ToString().Trim()) / 20));
                                }
                                break;
                            case "KUKGRD": //건강보험등급
                                if (oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Length > 0)
                                {
                                    oDS_PH_PY001A.SetValue("U_KUKGRD", 0, oForm.Items.Item(pVal.ItemUID).Specific.String.ToString("000"));
                                }
                                break;
                            case "BUYN20":
                            case "BUYN60": //부양가족20세이하 , 부양가족60세이상
                                oDS_PH_PY001A.SetValue("U_BUYNSU", 0, Convert.ToString(Convert.ToInt16(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) + Convert.ToInt16(oDS_PH_PY001A.GetValue("U_BUYN60", 0))));
                                break;
                            case "DAGYSU": //다자녀
                                if (Convert.ToInt16(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) < Convert.ToInt16(oDS_PH_PY001A.GetValue("U_DAGYSU", 0)))
                                {
                                    oDS_PH_PY001A.SetValue("U_BUYN20", 0, Convert.ToString(Convert.ToInt16(oDS_PH_PY001A.GetValue("U_DAGYSU", 0))));
                                    oDS_PH_PY001A.SetValue("U_BUYNSU", 0, Convert.ToString(Convert.ToInt16(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) + Convert.ToInt16(oDS_PH_PY001A.GetValue("U_BUYN60", 0)))); //부양가족수
                                }
                                break;
                            case "GBHAMT": //고용보험보수월액
                                if (Convert.ToDouble(oDS_PH_PY001A.GetValue("U_GBHAMT", 0)) == 0)
                                {
                                    oDS_PH_PY001A.SetValue("U_GBHSEL", 0, "N");
                                }
                                else
                                {
                                    oDS_PH_PY001A.SetValue("U_GBHSEL", 0, "Y");
                                }
                                break;
                            case "Mat1": //매트릭스
                                if (pVal.ColUID == "FILD01")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat2":
                                if (pVal.ColUID == "FILD01")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat3":
                                if (pVal.ColUID == "FamNam")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat4":
                                if (pVal.ColUID == "employer")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat4.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat5":
                                if (pVal.ColUID == "LCNumber")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat5.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat6":
                                if (pVal.ColUID == "appNum")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat6.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat7":
                                if (pVal.ColUID == "InDate")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat7.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat8":
                                if (pVal.ColUID == "School")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat8.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat9":
                                if (pVal.ColUID == "major")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat9.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat10":
                                if (pVal.ColUID == "Basis")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat10.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat11":
                                if (pVal.ColUID == "Passport")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat11.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat12":
                                if (pVal.ColUID == "Period")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat12.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat13":
                                if (pVal.ColUID == "Person")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat13.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat14":
                                if (pVal.ColUID == "injury")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat14.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                            case "Mat15":
                                if (pVal.ColUID == "Car")
                                {
                                    PH_PY001_AddMatrixRow();
                                    oMat15.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("가족사항 - 잘못된 주민번호입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("잘못된 주민번호입니다. : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                
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
                    oMat1.LoadFromDataSource();
                    oMat2.LoadFromDataSource();
                    oMat3.LoadFromDataSource();
                    oMat4.LoadFromDataSource();
                    oMat5.LoadFromDataSource();
                    oMat6.LoadFromDataSource();
                    oMat7.LoadFromDataSource();
                    oMat8.LoadFromDataSource();
                    oMat9.LoadFromDataSource();
                    oMat10.LoadFromDataSource();
                    oMat11.LoadFromDataSource();
                    oMat12.LoadFromDataSource();
                    oMat13.LoadFromDataSource();
                    oMat14.LoadFromDataSource();
                    oMat15.LoadFromDataSource();
                    PH_PY001_AddMatrixRow();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001C);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001D);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001E);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001F);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001G);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001J);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001K);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001N);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001P);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001Q);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY001R);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat3);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat4);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat5);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat6);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat7);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat8);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat9);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat10);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat11);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat12);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat13);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat14);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat15);
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
        /// FORM_RESIZE 이벤트
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
                    oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10;
                    oForm.Items.Item("79").Height = oForm.Items.Item("80").Height;

                    oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16;
                    oForm.Items.Item("77").Height = oForm.Items.Item("78").Height;

                    oForm.Items.Item("8").Width = oForm.Items.Item("Mat2").Left + oForm.Items.Item("Mat2").Width + 5;
                    oForm.Items.Item("8").Height = oForm.Items.Item("1").Top - 80;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 행삭제(사용자 메소드로 구현)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="oMat">매트릭스 이름</param>
        /// <param name="DBData">DB데이터소스</param>
        /// <param name="CheckField">데이터 체크 필드명</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;

            try
            {
                if (oLastColRow > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while (i <= DBData.Size - 1)
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord(i);
                                    i = 0;
                                }
                                else
                                {
                                    i += 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY001A", "Code"); //접속자 권한에 따른 사업장
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY001_FormItemEnabled();
                            PH_PY001_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY001_FormItemEnabled();
                            PH_PY001_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY001_FormItemEnabled();
                            PH_PY001_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY001_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat1, oDS_PH_PY001B, "U_FILD01"); //[MAT1] 급여 - 수당
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat2, oDS_PH_PY001C, "U_FILD01"); //[MAT2] 급여 - 공제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat3, oDS_PH_PY001D, "U_FamNam"); //[MAT3] 가족사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat4, oDS_PH_PY001E, "U_employer"); //[Mat4] 경력사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat5, oDS_PH_PY001F, "U_License"); //[Mat5] 자격사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat6, oDS_PH_PY001G, "U_appNum"); //[Mat6] 발령사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat7, oDS_PH_PY001H, "U_Club"); //[Mat7] 동호회
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat8, oDS_PH_PY001J, "U_School"); //[Mat8] 학력사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat9, oDS_PH_PY001K, "U_major"); //[Mat9] 교육사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat10, oDS_PH_PY001L, "U_Basis"); //[Mat10] 상벌사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat11, oDS_PH_PY001M, "U_Passport"); //[Mat11] 여권사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat12, oDS_PH_PY001N, "U_fromDate"); //[Mat12] 노조이력
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat13, oDS_PH_PY001P, "U_Person"); //[Mat13] 경조사항
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat14, oDS_PH_PY001Q, "U_fromDate"); //[Mat14] 사고현황
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat15, oDS_PH_PY001R, "U_CarNum"); //[Mat15] 차량현황
                            PH_PY001_AddMatrixRow();
                            break;
                    }
                }
            }
            catch(Exception ex)
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
            int i;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33

                            //부서
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Seq]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '1' ";
                            sQry += "             AND U_Char2 = '" + oDS_PH_PY001A.GetValue("U_CLTCOD", 0).ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");
                            oForm.Items.Item("TeamCode").DisplayDesc = true;

                            //담당
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Seq]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '2' ";
                            sQry += "             AND U_Char2 = '" + oDS_PH_PY001A.GetValue("U_CLTCOD", 0).ToString().Trim() + "'";
                            sQry += "             AND U_Char1 = '" + oDS_PH_PY001A.GetValue("U_TeamCode", 0).ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");
                            oForm.Items.Item("RspCode").DisplayDesc = true;

                            //반 (담당에 따른 반변경)
                            if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT      '' AS [Code], ";
                            sQry += "             '' AS [Name],";
                            sQry += "             -1 AS [Seq]";
                            sQry += " UNION ALL";
                            sQry += " SELECT      U_Code AS [Code], ";
                            sQry += "             U_CodeNm AS [Name],";
                            sQry += "             U_Seq AS [Seq]";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = '9' ";
                            sQry += "             AND U_Char3 = '" + oDS_PH_PY001A.GetValue("U_CLTCOD", 0).ToString().Trim() + "'";
                            sQry += "             AND U_Char2 = '" + oDS_PH_PY001A.GetValue("U_TeamCode", 0).ToString().Trim() + "'";
                            sQry += "             AND U_Char1 = '" + oDS_PH_PY001A.GetValue("U_RspCode", 0).ToString().Trim() + "'";
                            sQry += "             AND U_UseYN = 'Y'";
                            sQry += " ORDER BY    Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "N");
                            oForm.Items.Item("ClsCode").DisplayDesc = true;

                            //근무조
                            if (oForm.Items.Item("GNMUJO").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("GNMUJO").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("GNMUJO").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT      U_Code,";
                            sQry += "             U_CodeNm";
                            sQry += " FROM        [@PS_HR200L] ";
                            sQry += " WHERE       Code = 'P155'";
                            sQry += "             AND U_Char1 = '" + oDS_PH_PY001A.GetValue("U_ShiftDat", 0).ToString().Trim() + "'";
                            sQry += " ORDER BY    U_Code";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "N");
                            oForm.Items.Item("GNMUJO").DisplayDesc = true;

                            switch (oDS_PH_PY001A.GetValue("U_PAYTYP", 0).ToString().Trim())
                            {
                                case "1":
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = true;
                                    break;
                                case "2":
                                    oForm.Items.Item("Btn02").Visible = true;
                                    oForm.Items.Item("HOBONG").Enabled = true;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                                case "3":
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                                case "4":
                                    oForm.Items.Item("Btn02").Visible = false;
                                    oForm.Items.Item("HOBONG").Enabled = false;
                                    oForm.Items.Item("YUNAMT").Enabled = false;
                                    break;
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            oForm.Items.Item("Code").Enabled = false;
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    case "Mat1":
                    case "Mat2":
                    case "Mat3":
                    case "Mat4":
                    case "Mat5":
                    case "Mat6":
                    case "Mat7":
                    case "Mat8":
                    case "Mat9":
                    case "Mat10":
                    case "Mat11":
                    case "Mat12":
                    case "Mat13":
                    case "Mat14":
                    case "Mat15":
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
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
