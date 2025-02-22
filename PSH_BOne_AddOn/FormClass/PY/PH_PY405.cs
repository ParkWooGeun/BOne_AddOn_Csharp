﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 의료비자료등록
    /// </summary>
    internal class PH_PY405 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY405;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY405L;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY405.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY405_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY405");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY405_CreateItems();
                PH_PY405_FormItemEnabled();
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY405_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY405");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY405");
                oDS_PH_PY405 = oForm.DataSources.DataTables.Item("PH_PY405");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oDS_PH_PY405L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("년도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("관계코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("관계명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("주민번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급처상호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("사업자번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급일자", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("의료증빙코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("의료증빙명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급금액(외)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("지급금액(국)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("건수", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("의료비구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("경로", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("장애", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("난임", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("특례", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY405").Columns.Add("미숙", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
                oForm.DataSources.UserDataSources.Item("Year").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                // 관계
                oForm.DataSources.UserDataSources.Add("rel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("rel").Specific.DataBind.SetBound(true, "", "rel");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("rel").Specific, "Y");
                oForm.Items.Item("rel").DisplayDesc = true;
                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 성명
                oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

                // 주민등록번호
                oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

                // 내.외국인코드
                oForm.DataSources.UserDataSources.Add("empdiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("empdiv").Specific.DataBind.SetBound(true, "", "empdiv");
                oForm.Items.Item("empdiv").Specific.ValidValues.Add("1", "내국인");
                oForm.Items.Item("empdiv").Specific.ValidValues.Add("9", "외국인");
                oForm.Items.Item("empdiv").DisplayDesc = true;
                oForm.Items.Item("empdiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 지급처상호
                oForm.DataSources.UserDataSources.Add("custnm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("custnm").Specific.DataBind.SetBound(true, "", "custnm");

                // 사업자등록번호
                oForm.DataSources.UserDataSources.Add("entno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("entno").Specific.DataBind.SetBound(true, "", "entno");

                // 지급일자
                oForm.DataSources.UserDataSources.Add("payymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("payymd").Specific.DataBind.SetBound(true, "", "payymd");

                // 의료증빙코드
                oForm.DataSources.UserDataSources.Add("gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("gubun").Specific.DataBind.SetBound(true, "", "gubun");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("1", "국세청장이제공하는의료비자료");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("2", "국민건강보험공단의의료비부담명세서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("3", "진료비계산서,약제비계산서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("4", "장기요양급여비용명세서");
                oForm.Items.Item("gubun").Specific.ValidValues.Add("5", "기타의료비영수증");
                oForm.Items.Item("gubun").DisplayDesc = true;
                oForm.Items.Item("gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 지급금액(국세청자료외)
                oForm.DataSources.UserDataSources.Add("medcex", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("medcex").Specific.DataBind.SetBound(true, "", "medcex");

                // 지급금액(국세청자료)
                oForm.DataSources.UserDataSources.Add("ntamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntamt").Specific.DataBind.SetBound(true, "", "ntamt");

                // 지급건수
                oForm.DataSources.UserDataSources.Add("cont", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("cont").Specific.DataBind.SetBound(true, "", "cont");

                // 의료비구분
                oForm.DataSources.UserDataSources.Add("mcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("mcode").Specific.DataBind.SetBound(true, "", "mcode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '72' AND U_Char3 = '10' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("mcode").Specific, "Y");
                oForm.Items.Item("mcode").DisplayDesc = true;
                oForm.Items.Item("mcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 경로여부
                oForm.DataSources.UserDataSources.Add("olddiv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("olddiv").Specific.DataBind.SetBound(true, "", "olddiv");
                oForm.Items.Item("olddiv").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("olddiv").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("olddiv").DisplayDesc = true;
                oForm.Items.Item("olddiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 장애여부
                oForm.DataSources.UserDataSources.Add("deform", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("deform").Specific.DataBind.SetBound(true, "", "deform");
                oForm.Items.Item("deform").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("deform").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("deform").DisplayDesc = true;
                oForm.Items.Item("deform").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 난임시술비여부
                oForm.DataSources.UserDataSources.Add("nanim", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("nanim").Specific.DataBind.SetBound(true, "", "nanim");
                oForm.Items.Item("nanim").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("nanim").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("nanim").DisplayDesc = true;
                oForm.Items.Item("nanim").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 건겅보험산정특례자여부
                oForm.DataSources.UserDataSources.Add("tukrae", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("tukrae").Specific.DataBind.SetBound(true, "", "tukrae");
                oForm.Items.Item("tukrae").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("tukrae").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("tukrae").DisplayDesc = true;
                oForm.Items.Item("tukrae").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 미숙아.선천성이상아여부
                oForm.DataSources.UserDataSources.Add("prebaby", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("prebaby").Specific.DataBind.SetBound(true, "", "prebaby");
                oForm.Items.Item("prebaby").Specific.ValidValues.Add("N", "N");
                oForm.Items.Item("prebaby").Specific.ValidValues.Add("Y", "Y");
                oForm.Items.Item("prebaby").DisplayDesc = true;
                oForm.Items.Item("prebaby").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY405_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                oForm.EnableMenu("1282", true);  // 문서추가
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                }

                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                oForm.Items.Item("juminno").Specific.Value = "";
                oForm.Items.Item("custnm").Specific.Value = "";
                oForm.Items.Item("entno").Specific.Value = "";
                oForm.Items.Item("payymd").Specific.Value = "";

                oForm.Items.Item("medcex").Specific.Value = 0;
                oForm.Items.Item("ntamt").Specific.Value = 0;
                oForm.Items.Item("cont").Specific.Value = 0;

                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("olddiv").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("deform").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("nanim").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("tukrae").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("prebaby").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("mcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;

                oForm.Items.Item("juminno").Enabled = true;
                oForm.Items.Item("custnm").Enabled = true;
                oForm.Items.Item("payymd").Enabled = true;
                oForm.Items.Item("entno").Enabled = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY405_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY405_DataFind
        /// </summary>
        private void PH_PY405_DataFind()
        {
            string sQry;

            try
            {
                oForm.Freeze(true);
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return;
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번이 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return;
                }
                PH_PY405_FormItemEnabled();

                sQry = "EXEC PH_PY405_01 '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                oDS_PH_PY405.ExecuteQuery(sQry);
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY405_SAVE
        /// </summary>
        private void PH_PY405_SAVE()
        {
            // 데이타 저장
            string sQry;
            string saup;
            string yyyy;
            string sabun;
            string kname;
            string juminno;
            string custnm;
            string payymd;
            string rel;
            string empdiv;
            string entno;
            string Gubun;
            string olddiv;
            string deform;
            string nanim;
            string tukrae;
            string prebaby;
            string mcode;
            double medcex;
            double ntamt;
            double cont;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                rel = oForm.Items.Item("rel").Specific.Value.ToString().Trim();
                kname = oForm.Items.Item("kname").Specific.Value.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.Value.ToString().Trim();
                empdiv = oForm.Items.Item("empdiv").Specific.Value.ToString().Trim();
                custnm = oForm.Items.Item("custnm").Specific.Value.ToString().Trim();
                entno = oForm.Items.Item("entno").Specific.Value.ToString().Trim();
                payymd = oForm.Items.Item("payymd").Specific.Value.ToString().Trim();
                Gubun = oForm.Items.Item("gubun").Specific.Value.ToString().Trim();
                medcex = Convert.ToDouble(oForm.Items.Item("medcex").Specific.Value.ToString().Trim());
                ntamt = Convert.ToDouble(oForm.Items.Item("ntamt").Specific.Value.ToString().Trim());
                cont = Convert.ToDouble(oForm.Items.Item("cont").Specific.Value.ToString().Trim());
                olddiv = oForm.Items.Item("olddiv").Specific.Value.ToString().Trim();
                deform = oForm.Items.Item("deform").Specific.Value.ToString().Trim();
                nanim = oForm.Items.Item("nanim").Specific.Value.ToString().Trim();
                tukrae = oForm.Items.Item("tukrae").Specific.Value.ToString().Trim();
                prebaby = oForm.Items.Item("prebaby").Specific.Value.ToString().Trim();
                mcode = oForm.Items.Item("mcode").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(yyyy))
                {
                    PSH_Globals.SBO_Application.MessageBox("년도가 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(saup))
                {
                    PSH_Globals.SBO_Application.MessageBox("사업장이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(sabun))
                {
                    PSH_Globals.SBO_Application.MessageBox("사번이 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(mcode))
                {
                    PSH_Globals.SBO_Application.MessageBox("의료비구분코드를 반드시 입력하세요. 확인바랍니다..");
                    return;
                }

                if (olddiv == "Y" && deform == "Y")
                {
                    PSH_Globals.SBO_Application.MessageBox("경로여부와 장애여부는 둘다'Y'일 수 없습니다. 확인바랍니다..");
                    return;
                }

                if (string.IsNullOrEmpty(juminno) || (medcex == 0 && ntamt == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                if (medcex != 0 && ntamt != 0)
                {
                    PSH_Globals.SBO_Application.MessageBox("국세청자료와 국세청자료외는 구분하여 별도로 입력 하십시요. 확인바랍니다..");
                    return;
                }

                if (medcex != 0)
                {
                    if (string.IsNullOrEmpty(entno))
                    {
                        PSH_Globals.SBO_Application.MessageBox("사업자등록번호를 확인바랍니다..");
                        return;
                    }
                    if (cont == 0)
                    {
                        PSH_Globals.SBO_Application.MessageBox("지급건수를 확인바랍니다..");
                        return;
                    }
                }

                sQry = " Select Count(*) From [p_seoymedhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    sQry = "Update [p_seoymedhis] set ";
                    sQry += "rel = '" + rel + "',";
                    sQry += "kname = '" + kname + "',";
                    sQry += "juminno = '" + juminno + "',";
                    sQry += "empdiv = '" + empdiv + "',";
                    sQry += "custnm = '" + custnm + "',";
                    sQry += "entno = '" + entno + "',";
                    sQry += "payymd = '" + payymd + "',";
                    sQry += "gubun = '" + Gubun + "',";
                    sQry += "medcex = " + medcex + ",";
                    sQry += "ntamt = " + ntamt + ",";
                    sQry += "cont = " + cont + ",";
                    sQry += "olddiv = '" + olddiv + "',";
                    sQry += "deform = '" + deform + "',";
                    sQry += "tukrae = '" + tukrae + "',";
                    sQry += "prebaby = '" + prebaby + "',";
                    sQry += "mcode = '" + mcode + "',";
                    sQry += "nanim = '" + nanim + "'";
                    sQry += " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY405_DataFind();
                }
                else
                {
                    // 신규
                    sQry = "INSERT INTO [p_seoymedhis]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "rel,";
                    sQry += "kname,";
                    sQry += "juminno,";
                    sQry += "empdiv,";
                    sQry += "custnm,";
                    sQry += "entno,";
                    sQry += "payymd,";
                    sQry += "gubun,";
                    sQry += "medcex,";
                    sQry += "ntamt,";
                    sQry += "cont,";
                    sQry += "olddiv,";
                    sQry += "deform,";
                    sQry += "nanim,";
                    sQry += "tukrae,";
                    sQry += "prebaby,";
                    sQry += "mcode,";
                    sQry += "mednm";
                    sQry += " ) ";
                    sQry += "VALUES(";
                    sQry += "'" + saup + "',";
                    sQry += "'" + yyyy + "',";
                    sQry += "'" + sabun + "',";
                    sQry += "'" + rel + "',";
                    sQry += "'" + kname + "',";
                    sQry += "'" + juminno + "',";
                    sQry += "'" + empdiv + "',";
                    sQry += "'" + custnm + "',";
                    sQry += "'" + entno + "',";
                    sQry += "'" + payymd + "',";
                    sQry += "'" + Gubun + "',";
                    sQry += medcex + ",";
                    sQry += ntamt + ",";
                    sQry += cont + ",";
                    sQry += "'" + olddiv + "',";
                    sQry += "'" + deform + "',";
                    sQry += "'" + nanim + "',";
                    sQry += "'" + tukrae + "',";
                    sQry += "'" + prebaby + "',";
                    sQry += "'" + mcode + "',";
                    sQry += "'" + "" + "'";
                    sQry += " ) ";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY405_DataFind();
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY405_Delete
        /// </summary>
        private void PH_PY405_Delete()
        {
            // 데이타 삭제
            string sQry;
            string saup;
            string yyyy;
            string sabun;
            string juminno;
            string custnm;
            string entno;
            string payymd;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value;
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.Value;
                custnm = oForm.Items.Item("custnm").Specific.Value;
                entno = oForm.Items.Item("entno").Specific.Value;
                payymd = oForm.Items.Item("payymd").Specific.Value;

                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    if (oDS_PH_PY405.Rows.Count > 0)
                    {
                        sQry = "Delete From [p_seoymedhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "' And juminno = '" + juminno + "' And custnm = '" + custnm + "' And payymd = '" + payymd + "' And entno = '" + entno + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY405_DataFind();
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY405_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

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
            string Result;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_ret") // 조회
                    {
                        PH_PY405_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY405_SAVE();
                        }
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY405_Delete();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        oForm.Items.Item("kname").Specific.Value = oMat01.Columns.Item("kname").Cells.Item(pVal.Row).Specific.Value;
                        oForm.Items.Item("juminno").Specific.Value = oMat01.Columns.Item("juminno").Cells.Item(pVal.Row).Specific.Value;
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            int i;
            string sQry;
            string CLTCOD;
            string MSTCOD;
            string yyyy;
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
                        if (pVal.ItemUID == "rel")
                        {
                            oMat01.Clear();
                            oDS_PH_PY405L.Clear();

                            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                            yyyy = oForm.Items.Item("Year").Specific.Value;

                            if (!string.IsNullOrEmpty(oForm.Items.Item("rel").Specific.Value))
                            {
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                            }

                            sQry = "Select Distinct kname, juminno, birthymd, relatenm = ( select U_CodeNm From[@PS_HR200L] WHERE Code = 'P121' AND U_Code = relate) ";
                            sQry += " From [p_seoybase]";
                            sQry += " Where saup = '" + CLTCOD + "'";
                            sQry += " and sabun = '" + MSTCOD + "'";
                            sQry += " and div In ('10','70') ";
                            sQry += " and relate = '" + oForm.Items.Item("rel").Specific.Value + "'";
                            sQry += " and yyyy = '" + yyyy + "'";

                            oRecordSet.DoQuery(sQry);

                            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                            {
                                if (i + 1 > oDS_PH_PY405L.Size)
                                {
                                    oDS_PH_PY405L.InsertRecord(i);
                                }

                                oMat01.AddRow();
                                oDS_PH_PY405L.Offset = i;

                                oDS_PH_PY405L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                oDS_PH_PY405L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("kname").Value.ToString().Trim());
                                oDS_PH_PY405L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("juminno").Value.ToString().Trim());
                                oDS_PH_PY405L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("birthymd").Value.ToString().Trim());
                                oDS_PH_PY405L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("relatenm").Value.ToString().Trim());
                                oRecordSet.MoveNext();
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                  
                            if (oRecordSet.RecordCount == 1)
                            {
                                oForm.Items.Item("kname").Specific.Value = oMat01.Columns.Item("kname").Cells.Item(1).Specific.Value;
                                oForm.Items.Item("juminno").Specific.Value = oMat01.Columns.Item("juminno").Cells.Item(1).Specific.Value;
                            }
                        }
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string sQry;
            string CLTCOD;
            string MSTCOD;
            string FullName;
            string rel;
            string kname;
            string yyyy;
            string juminno;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                            case "MSTCOD":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                                sQry = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += " and Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "FullName":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                FullName = oForm.Items.Item("FullName").Specific.Value;

                                sQry = "Select Code,";
                                sQry += " FullName = U_FullName,";
                                sQry += " TeamName = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '1'";
                                sQry += " And U_Code = U_TeamCode),''),";
                                sQry += " RspName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '2'";
                                sQry += " And U_Code = U_RspCode),''),";
                                sQry += " ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += " From [@PS_HR200L]";
                                sQry += " WHERE Code = '9'";
                                sQry += " And U_Code  = U_ClsCode";
                                sQry += " And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += " And U_status <> '5'"; // 퇴사자 제외
                                sQry += " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //oForm.Items("MSTCOD").Specific.Value = oRecordSet.Fields("Code").Value
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "kname":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                                rel = oForm.Items.Item("rel").Specific.Value;
                                kname = oForm.Items.Item("kname").Specific.Value;
                                yyyy = oForm.Items.Item("Year").Specific.Value;

                                oForm.Items.Item("juminno").Specific.Value = "";

                                sQry = "Select Distinct juminno ";
                                sQry += " From [p_seoybase]";
                                sQry += " Where saup = '" + CLTCOD + "'";
                                sQry += " and sabun = '" + MSTCOD + "'";
                                sQry += " and relate = '" + rel + "'";
                                sQry += " and kname = '" + kname + "'";
                                sQry += " and yyyy = '" + yyyy + "'";

                                oRecordSet.DoQuery(sQry);

                                juminno = oRecordSet.Fields.Item("juminno").Value;
                                if (!string.IsNullOrEmpty(juminno))
                                {
                                    oForm.Items.Item("juminno").Specific.Value = juminno;

                                    if (rel != "01")
                                    {
                                        // 65세 경로우대 의료비 체크
                                        sQry = "select Cnt = Count(*) from p_seoybase a ";
                                        sQry += " Where a.yyyy = '" + yyyy + "'";
                                        sQry += " and datediff(yy, Left(a.birthymd,4) + '1231', '" + yyyy + "1231'" + " ) >= 65";
                                        sQry += " And a.juminno = '" + juminno + "'";
                                        oRecordSet.DoQuery(sQry);

                                        if (oRecordSet.Fields.Item("Cnt").Value > 0)
                                        {
                                            oForm.Items.Item("olddiv").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        else
                                        {
                                            oForm.Items.Item("olddiv").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }

                                        // 장애자인경우
                                        if (oForm.Items.Item("olddiv").Specific.Value.ToString().Trim() == "N")  //경로가 Check된경우는 제외
                                        {
                                            sQry = " Select Cnt = Count(*) From p_seoybase ";
                                            sQry += " Where yyyy = '" + yyyy + "'";
                                            sQry += " and div = '20' and target = '220'";
                                            sQry += " And juminno = '" + juminno + "'";
                                            oRecordSet.DoQuery(sQry);

                                            if (oRecordSet.Fields.Item("Cnt").Value > 0)
                                            {
                                                oForm.Items.Item("deform").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                            }
                                            else
                                            {
                                                oForm.Items.Item("deform").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        oForm.Items.Item("olddiv").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("deform").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("nanim").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("tukrae").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("prebaby").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.SetStatusBarMessage("기본공제대상자가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    return;
                                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                string sQry = string.Empty;
                string Param01;
                string Param02;
                string Param03;
                string Param04;
                string Param05;
                string Param06;
                string Param07;
                SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);
                            Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            Param02 = oDS_PH_PY405.Columns.Item("연도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY405.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY405.Columns.Item("주민번호").Cells.Item(pVal.Row).Value;
                            Param05 = oDS_PH_PY405.Columns.Item("지급처상호").Cells.Item(pVal.Row).Value;
                            Param06 = oDS_PH_PY405.Columns.Item("지급일자").Cells.Item(pVal.Row).Value;
                            Param07 = oDS_PH_PY405.Columns.Item("사업자번호").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY405_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "'";
                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.RecordCount == 0)
                            {
                                oForm.Items.Item("kname").Specific.Value = "";
                                oForm.Items.Item("juminno").Specific.Value = "";
                                oForm.Items.Item("custnm").Specific.Value = "";
                                oForm.Items.Item("entno").Specific.Value = "";
                                oForm.Items.Item("payymd").Specific.Value = "";

                                oForm.Items.Item("medcex").Specific.Value = 0;
                                oForm.Items.Item("ntamt").Specific.Value = 0;
                                oForm.Items.Item("cont").Specific.Value = 0;

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {
                                oForm.Items.Item("rel").Specific.Select(oRecordSet.Fields.Item("rel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value;
                                oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value;

                                oForm.Items.Item("empdiv").Specific.Select(oRecordSet.Fields.Item("empdiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("custnm").Value = oRecordSet.Fields.Item("custnm").Value;
                                oForm.DataSources.UserDataSources.Item("entno").Value = oRecordSet.Fields.Item("entno").Value;
                                oForm.DataSources.UserDataSources.Item("payymd").Value = oRecordSet.Fields.Item("payymd").Value;

                                oForm.Items.Item("gubun").Specific.Select(oRecordSet.Fields.Item("gubun").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("medcex").Value = oRecordSet.Fields.Item("medcex").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ntamt").Value = oRecordSet.Fields.Item("ntamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("cont").Value = oRecordSet.Fields.Item("cont").Value.ToString();

                                oForm.Items.Item("olddiv").Specific.Select(oRecordSet.Fields.Item("olddiv").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("deform").Specific.Select(oRecordSet.Fields.Item("deform").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("nanim").Specific.Select(oRecordSet.Fields.Item("nanim").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("tukrae").Specific.Select(oRecordSet.Fields.Item("tukrae").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("prebaby").Specific.Select(oRecordSet.Fields.Item("prebaby").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("mcode").Specific.Select(oRecordSet.Fields.Item("mcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                oForm.Items.Item("CLTCOD").Enabled = false;
                                oForm.Items.Item("Year").Enabled = false;
                                oForm.Items.Item("MSTCOD").Enabled = false;

                                oForm.Items.Item("juminno").Enabled = false;
                                oForm.Items.Item("custnm").Enabled = false;
                                oForm.Items.Item("payymd").Enabled = false;
                                oForm.Items.Item("entno").Enabled = false;
                            }
                        }
                    }

                    if (pVal.ItemUID == "Mat01")
                    {
                        oForm.Items.Item("kname").Specific.Value = oMat01.Columns.Item("kname").Cells.Item(pVal.Row).Specific.Value;
                        oForm.Items.Item("juminno").Specific.Value = oMat01.Columns.Item("juminno").Cells.Item(pVal.Row).Specific.Value;
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
                oForm.Freeze(false);
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY405);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY405L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            PH_PY405_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY405_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY405_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY405_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
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
    }
}
