using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 퇴직금기준설정
    /// </summary>
    internal class PH_PY114 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.Matrix oMat2;
        private SAPbouiCOM.DBDataSource oDS_PH_PY114A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY114B;
        private SAPbouiCOM.DBDataSource oDS_PH_PY114C;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string[] CSUCOD = new string[10];
        private string[] CSUNAM = new string[10];
        private string[] SILCOD = new string[28];
        private string[] SILCUN = new string[28];

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY114.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY114_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY114");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";
                
                oForm.Freeze(true);
                PH_PY114_CreateItems();
                PH_PY114_EnableMenus();
                PH_PY114_SetDocument(oFormDocEntry01);
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
        private void PH_PY114_CreateItems()
        {
            string sQry;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY114A = oForm.DataSources.DBDataSources.Item("@PH_PY114A");                ////헤더
                oDS_PH_PY114B = oForm.DataSources.DBDataSources.Item("@PH_PY114B");                ////라인
                oDS_PH_PY114C = oForm.DataSources.DBDataSources.Item("@PH_PY114C");                ////라인

                oMat1 = oForm.Items.Item("Mat1").Specific;
                oForm.Items.Item("Mat1").Specific.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oForm.Items.Item("Mat1").Specific.AutoResizeColumns();

                oMat2 = oForm.Items.Item("Mat2").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                oForm.PaneLevel = 1;

                //UserDataSources
                var _with1 = oForm.DataSources.UserDataSources;
                _with1.Add("MEDNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("KUKNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("JUNNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("BH1NAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("BH2NAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("CSUNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                _with1.Add("GONNAM", SAPbouiCOM.BoDataType.dt_LONG_TEXT);

                oForm.Items.Item("MEDNAM").Specific.DataBind.SetBound(true, "", "MEDNAM");
                oForm.Items.Item("KUKNAM").Specific.DataBind.SetBound(true, "", "KUKNAM");
                oForm.Items.Item("JUNNAM").Specific.DataBind.SetBound(true, "", "JUNNAM");
                oForm.Items.Item("BH1NAM").Specific.DataBind.SetBound(true, "", "BH1NAM");
                oForm.Items.Item("BH2NAM").Specific.DataBind.SetBound(true, "", "BH2NAM");
                oForm.Items.Item("CSUNAM").Specific.DataBind.SetBound(true, "", "CSUNAM");
                oForm.Items.Item("GONNAM").Specific.DataBind.SetBound(true, "", "GONNAM");


                //1.평균임금산정방법
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("M1", "월단위(3개월/월중도퇴사자 4개월)");
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("M2", "월단위(3개월/월말퇴사:~당월, 월중퇴사:~전월)");
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("M3", "월단위(3개월기준급여)");
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("D1", "일단위(퇴사월말일+2개월말일:07.11.13~08.02.11)");
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("D2", "일단위(퇴사일-이전90일까지: 07.11.14~08.02.11)");
                oForm.Items.Item("FIXTYP").Specific.ValidValues.Add("D3", "일단위(퇴사일-2개월퇴사전일:07.11.12~08.02.11)");

                //2.일할 끝전처리
                oForm.Items.Item("ROUNDT").Specific.ValidValues.Add("R", "반올림");
                oForm.Items.Item("ROUNDT").Specific.ValidValues.Add("F", "절사");
                oForm.Items.Item("ROUNDT").Specific.ValidValues.Add("C", "절상");

                oMat1.Columns.Item("ROUNDT").ValidValues.Add("R", "반올림");
                oMat1.Columns.Item("ROUNDT").ValidValues.Add("F", "절사");
                oMat1.Columns.Item("ROUNDT").ValidValues.Add("C", "절상");

                //3.단위
                oForm.Items.Item("LENGTH").Specific.ValidValues.Add("1", "  원");
                oForm.Items.Item("LENGTH").Specific.ValidValues.Add("10", "십원");
                oForm.Items.Item("LENGTH").Specific.ValidValues.Add("100", "백원");
                oForm.Items.Item("LENGTH").Specific.ValidValues.Add("1000", "천원");

                oMat1.Columns.Item("LENGTH").ValidValues.Add("1", "  원");
                oMat1.Columns.Item("LENGTH").ValidValues.Add("10", "십원");
                oMat1.Columns.Item("LENGTH").ValidValues.Add("100", "백원");
                oMat1.Columns.Item("LENGTH").ValidValues.Add("1000", "천원");

                //4.옵션버튼(비정기상여포함여부)
                oForm.Items.Item("Opt11").Specific.GroupWith("Opt12");
                oForm.Items.Item("Opt11").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_RETCHK");
                
                oForm.Items.Item("Opt12").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_RETCHK");
                oForm.Items.Item("Opt12").Specific.GroupWith("Opt11");

                oForm.Items.Item("Opt11").Specific.Selected = true;
                //5.옵션버튼(1년미만자 포함여부)
                oForm.Items.Item("Opt21").Specific.GroupWith("Opt22");
                oForm.Items.Item("Opt21").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_1YYCHK");

                oForm.Items.Item("Opt22").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_1YYCHK");
                oForm.Items.Item("Opt22").Specific.GroupWith("Opt21");
                oForm.Items.Item("Opt22").Specific.Selected = true;

                //6.옵션버튼(근속년수 기준)
                oForm.Items.Item("Opt31").Specific.GroupWith("Opt32");
                oForm.Items.Item("Opt31").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_GNSGBN");

                oForm.Items.Item("Opt32").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_GNSGBN");
                oForm.Items.Item("Opt32").Specific.GroupWith("Opt31");
                oForm.Items.Item("Opt31").Specific.Selected = true;

                //7.옵션버튼(분개생성 방법)
                oForm.Items.Item("Opt41").Specific.GroupWith("Opt42");
                oForm.Items.Item("Opt41").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_BILCHK");

                oForm.Items.Item("Opt42").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_BILCHK");
                oForm.Items.Item("Opt42").Specific.GroupWith("Opt41");
                oForm.Items.Item("Opt41").Specific.Selected = true;

                //8.옵션버튼(퇴사월 연차수당 처리)
                oForm.Items.Item("Opt51").Specific.GroupWith("Opt52");
                oForm.Items.Item("Opt51").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_RETYCH");
                oForm.Items.Item("Opt52").Specific.DataBind.SetBound(true, "@PH_PY114A", "U_RETYCH");
                oForm.Items.Item("Opt52").Specific.GroupWith("Opt51");
                oForm.Items.Item("Opt51").Specific.Selected = true;

                //초기값 셋팅
                for (i = 1; i <= 9; i++)
                {
                    CSUCOD[i] = "A0" + i;
                }
                CSUNAM[1] = "평균급여";
                CSUNAM[2] = "평균상여";
                CSUNAM[3] = "평균연차";
                CSUNAM[4] = "평균체력단련비";
                CSUNAM[5] = "평균임금";
                CSUNAM[6] = "연퇴직금";
                CSUNAM[7] = "월퇴직금";
                CSUNAM[8] = "일퇴직금";
                CSUNAM[9] = "퇴직금계";


                //M1 월단위계산수식
                SILCOD[1] = "P01";
                SILCUN[1] = "총급여";
                SILCOD[2] = "P02  / 12 * 3";
                SILCUN[2] = "총상여 / 12 * 3";
                SILCOD[3] = "P03  / 12 * 3";
                SILCUN[3] = "총연차 / 12 * 3";
                SILCOD[4] = "P04  / 12 * 3";
                SILCUN[4] = "총체력단련비 / 12 * 3";
                SILCOD[5] = "( A01 + A02 + A03 + A04 ) / X01";
                SILCUN[5] = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수";
                SILCOD[6] = "";
                SILCUN[6] = "";
                SILCOD[7] = "";
                SILCUN[7] = "";
                SILCOD[8] = "A05 * 30 * X02 / 365";
                SILCUN[8] = "평균임금*30*총근속일수/365";
                SILCOD[9] = "A06 + A07 + A08";
                SILCUN[9] = "연퇴직금+월퇴직금+일퇴직금";
                //D1 일단위계산수식
                SILCOD[10] = "P01";
                SILCUN[10] = "총급여";
                SILCOD[11] = "P02  / 12 * 3";
                SILCUN[11] = "총상여 / 12 * 3";
                SILCOD[12] = "P03  / 12 * 3";
                SILCUN[12] = "총연차 / 12 * 3";
                SILCOD[13] = "P04  / 12 * 3";
                SILCUN[13] = "총체력단련비 / 12 * 3";
                SILCOD[14] = "( A01 + A02 + A03 + A04 ) / X01";
                SILCUN[14] = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수수";
                SILCOD[15] = "";
                SILCUN[15] = "";
                SILCOD[16] = "";
                SILCUN[16] = "";
                SILCOD[17] = "A05 * 30 * X02 / 365";
                SILCUN[17] = "평균임금*30*총근속일수/365";
                SILCOD[18] = "A06 + A07 + A08";
                SILCUN[18] = "연퇴직금+월퇴직금+일퇴직금";
                //D2 일단위계산수식
                SILCOD[19] = "P01";
                SILCUN[19] = "총급여";
                SILCOD[20] = "P02  / 12 * 3";
                SILCUN[20] = "총상여 / 12 * 3";
                SILCOD[21] = "P03  / 12 * 3";
                SILCUN[21] = "총연차 / 12 * 3";
                SILCOD[22] = "P04  / 12 * 3";
                SILCUN[22] = "총체력단련비 / 12 * 3";
                SILCOD[23] = "( A01 + A02 + A03 + A04 ) / X01";
                SILCUN[23] = "(평균급여+평균상여+평균연차+평균체력단련비) / 급여기간일수";
                SILCOD[24] = "";
                SILCUN[24] = "";
                SILCOD[25] = "";
                SILCUN[25] = "";
                SILCOD[26] = "A05 * 30 * X02 / 365";
                SILCUN[26] = "평균임금*30*총근속일수/365";
                SILCOD[27] = "A06 + A07 + A08";
                SILCUN[27] = "연퇴직금+월퇴직금+일퇴직금";

                //Folder /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                oForm.DataSources.UserDataSources.Add("FolderDS1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                for (i = 1; i <= 3; i++)
                {
                    oForm.Items.Item("Folder" + i).Specific.DataBind.SetBound(true, "", "FolderDS1");
                    if (i == 1)
                    {
                        oForm.Items.Item("Folder" + i).Specific.Select();
                    }
                    else
                    {
                        oForm.Items.Item("Folder" + i).Specific.GroupWith("Folder" + (i - 1));
                    }
                }

                oForm.Items.Item("Folder1").AffectsFormMode = false;
                oForm.Items.Item("Folder2").AffectsFormMode = false;
                oForm.Items.Item("Folder3").AffectsFormMode = false;
                oForm.Items.Item("Folder1").Enabled = true;
                oForm.Items.Item("Folder2").Enabled = true;
                oForm.Items.Item("Folder3").Enabled = true;

                //직위
                sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID";
                oRecordSet.DoQuery(sQry);
                while (!(oRecordSet.EoF))
                {
                    oMat2.Columns.Item("MSTSTP").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }

                //귀속연월
                oForm.Items.Item("Code").Specific.Value = DateTime.Now.Year;

                //누진계산사용유무
                oForm.Items.Item("RATUSE").Specific.ValidValues.Add("N", "미사용");
                oForm.Items.Item("RATUSE").Specific.ValidValues.Add("Y", "사용");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY114_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY114_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY114_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY114_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY114_FormItemEnabled();
                    PH_PY114_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY114_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY114_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY114_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Opt11").Specific.Selected = true;
                    oForm.Items.Item("Opt22").Specific.Selected = true;
                    oForm.Items.Item("Opt31").Specific.Selected = true;
                    oForm.Items.Item("Opt42").Specific.Selected = true;
                    oForm.Items.Item("Opt51").Specific.Selected = true;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Opt11").Specific.Selected = true;
                    oForm.Items.Item("Opt22").Specific.Selected = true;
                    oForm.Items.Item("Opt31").Specific.Selected = true;
                    oForm.Items.Item("Opt42").Specific.Selected = true;
                    oForm.Items.Item("Opt51").Specific.Selected = true;

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY114_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private bool PH_PY114_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            int k;
            string DocNum;
            string Chk_Data;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY114A.GetValue("Code", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                if (string.IsNullOrEmpty(oDS_PH_PY114A.GetValue("U_FIXTYP", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("평균임금 산정방법은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                if (codeHelpClass.Left(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) == "D" & string.IsNullOrEmpty(oDS_PH_PY114A.GetValue("U_ROUNDT", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("일할 끝전처리방법을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                if (codeHelpClass.Left(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) == "D" & string.IsNullOrEmpty(oDS_PH_PY114A.GetValue("U_LENGTH", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("일할 단위를 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                DocNum = Exist_YN(oDS_PH_PY114A.GetValue("Code", 0));
                if (!string.IsNullOrEmpty(DocNum.ToString().Trim()) & oDS_PH_PY114A.GetValue("Code", 0).ToString().Trim() != DocNum.ToString().Trim())
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기존의 데이터가 있습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
                }
                //라인체크
                oMat1.FlushToDataSource();
                oMat2.FlushToDataSource();

                if (oMat1.RowCount == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
                }
                for (i = 0; i <= oMat1.VisualRowCount - 2; i++)
                {
                    oDS_PH_PY114B.Offset = i;
                    if (string.IsNullOrEmpty(oDS_PH_PY114B.GetValue("U_CSUCOD", i).ToString().Trim()))
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        oMat1.Columns.Item("Col4").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return functionReturnValue;
                    }
                    else
                    {
                        Chk_Data = oDS_PH_PY114B.GetValue("U_CSUCOD", i).ToString().Trim();
                        for (k = i + 1; k <= oMat1.VisualRowCount - 2; k++)
                        {
                            oDS_PH_PY114B.Offset = k;
                            if (Chk_Data.ToString().Trim() == oDS_PH_PY114B.GetValue("U_CSUCOD", k).ToString().Trim())
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oMat1.Columns.Item("Col1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }
                //Mat2
                if (oMat2.RowCount > 1)
                {
                    if (oDS_PH_PY114A.GetValue("U_RATUSE", 0) == "Y")
                    {
                        for (i = 0; i <= oMat2.VisualRowCount - 2; i++)
                        {
                            oDS_PH_PY114C.Offset = i;
                            if (string.IsNullOrEmpty(oDS_PH_PY114C.GetValue("U_MSTSTP", i).ToString().Trim()))
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("직위는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oMat2.Columns.Item("Col2").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                            else
                            {
                                Chk_Data = oDS_PH_PY114C.GetValue("U_MSTSTP", i).ToString().Trim();
                                for (k = i + 1; k <= oMat2.VisualRowCount - 2; k++)
                                {
                                    oDS_PH_PY114C.Offset = k;
                                    if (Chk_Data.ToString().Trim() == oDS_PH_PY114C.GetValue("U_MSTSTP", k).ToString().Trim())
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("임원누진 내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        oMat2.Columns.Item("Col2").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        return functionReturnValue;
                                    }
                                }
                            }
                        }
                    }
                    if (string.IsNullOrEmpty(oDS_PH_PY114C.GetValue("U_MSTSTP", oDS_PH_PY114C.Size - 1).ToString().Trim()))
                    {
                        oDS_PH_PY114C.RemoveRecord(oDS_PH_PY114C.Size - 1);
                    }
                }

                oMat1.LoadFromDataSource();
                oMat2.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY114_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY114_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            int ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY114A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    ErrNumm = 1;
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
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY114_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY114'", "");
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY114_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY114_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                //Mat2
                oMat2.FlushToDataSource();
                oRow = oMat2.VisualRowCount;

                if (oMat2.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY114C.GetValue("U_MSTSTP", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY114C.Size <= oMat2.VisualRowCount)
                        {
                            oDS_PH_PY114C.InsertRecord((oRow));
                        }
                        oDS_PH_PY114C.Offset = oRow;
                        oDS_PH_PY114C.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY114C.SetValue("U_MSTSTP", oRow, "");
                        oDS_PH_PY114C.SetValue("U_ADDRAT", oRow, "0");
                        oMat2.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY114C.Offset = oRow - 1;
                        oDS_PH_PY114C.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY114C.SetValue("U_MSTSTP", oRow - 1, "");
                        oDS_PH_PY114C.SetValue("U_ADDRAT", oRow - 1, "0");
                        oMat2.LoadFromDataSource();
                    }
                }
                else if (oMat2.VisualRowCount == 0)
                {
                    oDS_PH_PY114C.Offset = oRow;
                    oDS_PH_PY114C.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY114C.SetValue("U_MSTSTP", oRow, "");
                    oDS_PH_PY114C.SetValue("U_ADDRAT", oRow, "0");
                    oMat2.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY114_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Display_PH_PY114B
        /// </summary>
        private void Display_PH_PY114B()
        {
            string FIXTYP;
            int i;
            int cnt;

            try
            {
                FIXTYP = oForm.Items.Item("FIXTYP").Specific.Selected.Value;
                if (PSH_Globals.SBO_Application.MessageBox("기존 퇴직금 계산식을 초기 셋팅용 계산식으로 변경하시겠습니까?", 2, "&Yes!", "&No") == 2)
                {
                    return;
                }
                cnt = oDS_PH_PY114B.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY114B.RemoveRecord(oDS_PH_PY114B.Size - 1);
                    }
                }
                else
                {
                    oMat1.LoadFromDataSource();
                }
                switch (FIXTYP.ToString().Trim())
                {
                    case "M1":
                        cnt = 1;
                        break;
                    case "D1":
                        cnt = 10;
                        break;
                    default:
                        cnt = 19;
                        break;
                }
                for (i = 0; i <= 7; i++)
                {
                    if (i + 1 > oDS_PH_PY114B.Size)
                    {
                        oDS_PH_PY114B.InsertRecord(i);
                    }
                    oDS_PH_PY114B.Offset = i;
                    oDS_PH_PY114B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY114B.SetValue("U_CSUCOD", i, CSUCOD[i + 1]);
                    oDS_PH_PY114B.SetValue("U_CSUNAM", i, CSUNAM[i + 1]);
                    oDS_PH_PY114B.SetValue("U_SILCOD", i, SILCOD[i + cnt]);
                    oDS_PH_PY114B.SetValue("U_SILCUN", i, SILCUN[i + cnt]);
                    oDS_PH_PY114B.SetValue("U_ROUNDT", i, "R");
                    oDS_PH_PY114B.SetValue("U_LENGTH", i, "1");
                }
                oMat1.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Display_PH_PY114B_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Account_Name
        /// </summary>
        private void Account_Name()
        {
            string CSUACC;
            string BH1ACC;
            string KUKACC;
            string MEDACC;
            string JUNACC;
            string BH2ACC;
            string GONACC;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                MEDACC = oDS_PH_PY114A.GetValue("U_MEDACC", 0);
                KUKACC = oDS_PH_PY114A.GetValue("U_KUKACC", 0);
                JUNACC = oDS_PH_PY114A.GetValue("U_JUNACC", 0);
                BH1ACC = oDS_PH_PY114A.GetValue("U_BH1ACC", 0);
                BH2ACC = oDS_PH_PY114A.GetValue("U_BH2ACC", 0);
                CSUACC = oDS_PH_PY114A.GetValue("U_CSUACC", 0);
                GONACC = oDS_PH_PY114A.GetValue("U_GONACC", 0);

                if (string.IsNullOrEmpty(MEDACC))
                {
                    oForm.DataSources.UserDataSources.Item("MEDNAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("MEDNAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + MEDACC + "'", "");
                }

                if (string.IsNullOrEmpty(KUKACC))
                {
                    oForm.DataSources.UserDataSources.Item("KUKNAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("KUKNAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + KUKACC + "'", "");
                }

                if (string.IsNullOrEmpty(JUNACC))
                {
                    oForm.DataSources.UserDataSources.Item("JUNNAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("JUNNAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + JUNACC + "'", "");
                }

                if (string.IsNullOrEmpty(BH1ACC))
                {
                    oForm.DataSources.UserDataSources.Item("BH1NAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("BH1NAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + BH1ACC + "'", "");
                }

                if (string.IsNullOrEmpty(BH2ACC))
                {
                    oForm.DataSources.UserDataSources.Item("BH2NAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("BH2NAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + BH2ACC + "'", "");
                }

                if (string.IsNullOrEmpty(CSUACC))
                {
                    oForm.DataSources.UserDataSources.Item("CSUNAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("CSUNAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + CSUACC + "'", "");
                }

                if (string.IsNullOrEmpty(GONACC))
                {
                    oForm.DataSources.UserDataSources.Item("GONNAM").ValueEx = "";
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("GONNAM").ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + GONACC + "'", "");
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Account_Name_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private string Exist_YN(string STDDAT)
        {
            string sQry;
            string functionReturnValue = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT Top 1 T1.CODE FROM [@PH_PY114A] T1 where code = '" + STDDAT.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);

                if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
                {
                    functionReturnValue = "";
                }
                else
                {
                    functionReturnValue = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Exist_YN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return functionReturnValue;
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY114_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PH_PY114_DataValidCheck() == false)
                                {
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Folder1")
                    {
                        oForm.PaneLevel = 1;
                    }
                    else if (pVal.ItemUID == "Folder2")
                    {
                        oForm.PaneLevel = 2;
                    }
                    else if (pVal.ItemUID == "Folder3")
                    {
                        oForm.PaneLevel = 3;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
                if (pVal.BeforeAction == true && pVal.CharPressed == 9 && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {

                    if (pVal.ItemUID == "JOBDAT")
                    {
                        if (oMat2.RowCount > 0)
                        {
                            oMat2.Columns.Item("Col2").Cells.Item(oMat1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            BubbleEvent = false;
                        }
                    }
                }
                if (pVal.BeforeAction == true && pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "MEDACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("MEDACC").Specific.Value + "'","") == true || string.IsNullOrEmpty(oForm.Items.Item("MEDACC").Specific.String))
                        {
                            oForm.Items.Item("MEDACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "KUKACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("KUKACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("KUKACC").Specific.String))
                        {
                            oForm.Items.Item("KUKACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "JUNACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("JUNACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("JUNACC").Specific.String))
                        {
                            oForm.Items.Item("JUNACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "BH1ACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("BH1ACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("BH1ACC").Specific.String))
                        {
                            oForm.Items.Item("BH1ACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "BH2ACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("BH2ACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("BH2ACC").Specific.String))
                        {
                            oForm.Items.Item("BH2ACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "CSUACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("CSUACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("CSUACC").Specific.String))
                        {
                            oForm.Items.Item("CSUACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "GONACC")
                    {
                        if (dataHelpClass.Value_ChkYn("OACT", "FORMATCODE", "'" + oForm.Items.Item("GONACC").Specific.Value + "'", "") == true || string.IsNullOrEmpty(oForm.Items.Item("GONACC").Specific.String))
                        {
                            oForm.Items.Item("GONACC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    if (pVal.ItemUID == "Code")
                    {

                    }
                    else if (pVal.ItemUID == "Mat2" && (pVal.ColUID == "Col1" || pVal.ColUID == "Col2"))
                    {
                        oMat2.FlushToDataSource();
                        oDS_PH_PY114C.Offset = pVal.Row - 1;
                        oMat2.SetLineData(pVal.Row);

                        if (pVal.Row == oMat2.RowCount && !string.IsNullOrEmpty(oDS_PH_PY114C.GetValue("U_MSTSTP", pVal.Row - 1).ToString().Trim()))
                        {
                            PH_PY114_AddMatrixRow();
                            oMat2.Columns.Item("Col2").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                    }
                    else if (pVal.ItemUID == "MEDACC" || pVal.ItemUID == "KUKACC" || pVal.ItemUID == "JUNACC" || pVal.ItemUID == "BH1ACC" || pVal.ItemUID == "BH2ACC" || pVal.ItemUID == "CSUACC" || pVal.ItemUID == "GONACC")
                    {

                        if (string.IsNullOrEmpty(oDS_PH_PY114A.GetValue("U_" + pVal.ItemUID, 0)))
                        {
                            oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = "";
                        }
                        else
                        {
                            oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = dataHelpClass.Get_ReData("ACCTNAME", "FORMATCODE", "OACT", "'" + oDS_PH_PY114A.GetValue("U_" + pVal.ItemUID, 0) + "'","");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    oForm.Items.Item("Rec2").Width = oForm.Items.Item("Mat1").Width + 20;
                    oForm.Items.Item("Rec2").Height = oForm.Items.Item("Mat1").Height + 10;

                    oForm.Items.Item("Rec1").Width = oForm.Items.Item("Rec2").Width - oForm.Items.Item("Rec1").Left + 20;

                    oForm.Items.Item("Rec3").Top = oForm.Items.Item("Rec2").Top + oForm.Items.Item("Rec2").Height + 22;
                    oForm.Items.Item("Rec3").Width = oForm.Items.Item("Rec2").Width;
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.ItemChanged == true && pVal.ItemUID == "FIXTYP")
                {
                    if (pVal.BeforeAction == false)
                    {
                        oForm.Items.Item("FIXTYP").Update();
                        Display_PH_PY114B();
                        if (codeHelpClass.Left(oDS_PH_PY114A.GetValue("U_FIXTYP", 0), 1) == "M")
                        {
                            oDS_PH_PY114A.SetValue("U_ROUNDT", 0, "");
                            oDS_PH_PY114A.SetValue("U_LENGTH", 0, "");
                            oForm.Items.Item("ROUNDT").Enabled = false;
                            oForm.Items.Item("LENGTH").Enabled = false;
                            oForm.Items.Item("ROUNDT").Update();
                            oForm.Items.Item("LENGTH").Update();
                        }
                        else
                        {
                            oForm.Items.Item("ROUNDT").Enabled = true;
                            oForm.Items.Item("LENGTH").Enabled = true;
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
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == false && pVal.ItemUID == "Mat2" && oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oMat1.LoadFromDataSource();
                    oMat2.LoadFromDataSource();
                    PH_PY114_AddMatrixRow();
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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                switch (pVal.ItemUID)
                {
                    case "Mat1":
                    case "Grid1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                        case "Mat1":
                        case "Grid1":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY114A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY114B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY114C);
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
            int i = 0;
            
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
                            PH_PY114_FormItemEnabled();
                            PH_PY114_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY114_FormItemEnabled();
                            PH_PY114_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY114_FormItemEnabled();
                            PH_PY114_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY114_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            //MAT1
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while (i <= oDS_PH_PY114B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY114B.GetValue("U_CSUCOD", i)))
                                    {
                                        oDS_PH_PY114B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY114B.Size; i++)
                                {
                                    oDS_PH_PY114B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            //MAT2
                            if (oMat2.RowCount != oMat2.VisualRowCount)
                            {
                                oMat2.FlushToDataSource();

                                while (i <= oDS_PH_PY114B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY114B.GetValue("U_MSTSTP", i)))
                                    {
                                        oDS_PH_PY114B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY114B.Size; i++)
                                {
                                    oDS_PH_PY114B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat2.LoadFromDataSource();
                            }
                            PH_PY114_AddMatrixRow();
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

