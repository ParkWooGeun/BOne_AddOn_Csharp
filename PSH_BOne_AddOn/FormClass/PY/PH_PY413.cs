using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 월세액.주택임차차입금자료 등록
    /// </summary>
    internal class PH_PY413 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY413;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY413.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY413_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY413");

                string strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY413_CreateItems();
                PH_PY413_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
                oForm.Visible = true;
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY413_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY413");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY413");
                oDS_PH_PY413 = oForm.DataSources.DataTables.Item("PH_PY413");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                // 담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                // 반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");
                
                // 라인번호
                oForm.DataSources.UserDataSources.Add("LineNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("LineNum").Specific.DataBind.SetBound(true, "", "LineNum");

                // 임대인성명
                oForm.DataSources.UserDataSources.Add("ws_name", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ws_name").Specific.DataBind.SetBound(true, "", "ws_name");

                // 임대인주민등록번호
                oForm.DataSources.UserDataSources.Add("ws_jumin", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ws_jumin").Specific.DataBind.SetBound(true, "", "ws_jumin");

                // 주택유형
                oForm.DataSources.UserDataSources.Add("ws_hcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ws_hcode").Specific.DataBind.SetBound(true, "", "ws_hcode");
                sQry = "SELECT U_Char1, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ws_hcode").Specific, "Y");
                oForm.Items.Item("ws_hcode").DisplayDesc = true;
                oForm.Items.Item("ws_hcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 주택계약면적
                oForm.DataSources.UserDataSources.Add("ws_hm", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ws_hm").Specific.DataBind.SetBound(true, "", "ws_hm");

                // 임대차계약서상주소
                oForm.DataSources.UserDataSources.Add("ws_addr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ws_addr").Specific.DataBind.SetBound(true, "", "ws_addr");

                // 임대차계약기간
                oForm.DataSources.UserDataSources.Add("ws_fymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_fymd").Specific.DataBind.SetBound(true, "", "ws_fymd");

                oForm.DataSources.UserDataSources.Add("ws_tymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ws_tymd").Specific.DataBind.SetBound(true, "", "ws_tymd");

                // 월세액
                oForm.DataSources.UserDataSources.Add("ws_mamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_mamt").Specific.DataBind.SetBound(true, "", "ws_mamt");

                // 공제금액
                oForm.DataSources.UserDataSources.Add("ws_gamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ws_gamt").Specific.DataBind.SetBound(true, "", "ws_gamt");

                // 금전소비대차 계약내용
                // 대주
                oForm.DataSources.UserDataSources.Add("dj_name", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("dj_name").Specific.DataBind.SetBound(true, "", "dj_name");

                // 대주 주민등록번호
                oForm.DataSources.UserDataSources.Add("dj_jumin", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("dj_jumin").Specific.DataBind.SetBound(true, "", "dj_jumin");

                // 금전소비대차계약기간
                oForm.DataSources.UserDataSources.Add("dj_fymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_fymd").Specific.DataBind.SetBound(true, "", "dj_fymd");

                oForm.DataSources.UserDataSources.Add("dj_tymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("dj_tymd").Specific.DataBind.SetBound(true, "", "dj_tymd");

                // 차임급이자율
                oForm.DataSources.UserDataSources.Add("dj_eja", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("dj_eja").Specific.DataBind.SetBound(true, "", "dj_eja");

                // 계
                oForm.DataSources.UserDataSources.Add("dj_tamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_tamt").Specific.DataBind.SetBound(true, "", "dj_tamt");
                // 원리금
                oForm.DataSources.UserDataSources.Add("dj_wamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_wamt").Specific.DataBind.SetBound(true, "", "dj_wamt");

                // 이자
                oForm.DataSources.UserDataSources.Add("dj_eamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_eamt").Specific.DataBind.SetBound(true, "", "dj_eamt");

                // 공제금액
                oForm.DataSources.UserDataSources.Add("dj_gamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("dj_gamt").Specific.DataBind.SetBound(true, "", "dj_gamt");

                // 임대차계약내용	
                // 임대인성명
                oForm.DataSources.UserDataSources.Add("ld_name", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ld_name").Specific.DataBind.SetBound(true, "", "ld_name");

                // 임대인주민등록번호
                oForm.DataSources.UserDataSources.Add("ld_jumin", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("ld_jumin").Specific.DataBind.SetBound(true, "", "ld_jumin");

                // 주택유형
                oForm.DataSources.UserDataSources.Add("ld_hcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ld_hcode").Specific.DataBind.SetBound(true, "", "ld_hcode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ld_hcode").Specific, "Y");
                oForm.Items.Item("ld_hcode").DisplayDesc = true;
                oForm.Items.Item("ld_hcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 주택계약면적
                oForm.DataSources.UserDataSources.Add("ld_hm", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("ld_hm").Specific.DataBind.SetBound(true, "", "ld_hm");

                // 임대차계약서상주소
                oForm.DataSources.UserDataSources.Add("ld_addr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ld_addr").Specific.DataBind.SetBound(true, "", "ld_addr");

                // 임대차계약기간
                oForm.DataSources.UserDataSources.Add("ld_fymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_fymd").Specific.DataBind.SetBound(true, "", "ld_fymd");

                oForm.DataSources.UserDataSources.Add("ld_tymd", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ld_tymd").Specific.DataBind.SetBound(true, "", "ld_tymd");

                //전세보증금
                oForm.DataSources.UserDataSources.Add("ld_bamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ld_bamt").Specific.DataBind.SetBound(true, "", "ld_bamt");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY413_FormItemEnabled()
        {
            try
            {
                oForm.EnableMenu("1282", true);  // 문서추가
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                }
                oForm.Items.Item("MSTCOD").Specific.Value = "";
                oForm.DataSources.UserDataSources.Item("FullName").Value = "";
                oForm.Items.Item("FullName").Specific.Value = "";
                oForm.Items.Item("TeamName").Specific.Value = "";
                oForm.Items.Item("RspName").Specific.Value = "";
                oForm.Items.Item("ClsName").Specific.Value = "";
                oForm.Items.Item("LineNum").Specific.Value = "";

                oForm.DataSources.UserDataSources.Item("ws_name").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_jumin").Value = "";
                oForm.Items.Item("ws_hcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("ws_hm").Value = "0";
                oForm.DataSources.UserDataSources.Item("ws_addr").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_fymd").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_tymd").Value = "";
                oForm.DataSources.UserDataSources.Item("ws_mamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("ws_gamt").Value = "0";

                oForm.DataSources.UserDataSources.Item("dj_name").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_jumin").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_fymd").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tymd").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_eja").Value = "";
                oForm.DataSources.UserDataSources.Item("dj_tamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("dj_wamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("dj_eamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("dj_gamt").Value = "0";

                oForm.DataSources.UserDataSources.Item("ld_name").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_jumin").Value = "";
                oForm.Items.Item("ld_hcode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("ld_hm").Value = "0";
                oForm.DataSources.UserDataSources.Item("ld_addr").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_fymd").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_tymd").Value = "";
                oForm.DataSources.UserDataSources.Item("ld_bamt").Value = "0";

                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PH_PY413_DataFind
        /// </summary>
        private void PH_PY413_DataFind()
        {
            int iRow;
            string sQry;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    errMessage = "년도가 없습니다. 확인바랍니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim()))
                {
                    errMessage = "사업장이 없습니다. 확인바랍니다.";
                    throw new Exception();
                }

                PH_PY413_FormItemEnabled();

                sQry = "EXEC PH_PY413_01 '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                oDS_PH_PY413.ExecuteQuery(sQry);
                iRow = oDS_PH_PY413.Rows.Count;
                oForm.Items.Item("LineNum").Specific.Value = "";
                oForm.DataSources.UserDataSources.Item("LineNum").Value = "";
                oForm.ActiveItem = "ws_name";
                oGrid1.AutoResizeColumns();
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
            }
        }

        /// <summary>
        /// PH_PY413_SAVE
        /// </summary>
        private void PH_PY413_SAVE()
        {
            // 데이타 저장
            string sQry;
            string seqn;
            string saup;
            string yyyy;
            string sabun;
            string ws_name;
            string ws_jumin;
            string ws_hcode;
            string ws_addr;
            string ws_fymd;
            string ws_tymd;
            string dj_name;
            string dj_jumin;
            string dj_fymd;
            string dj_tymd;
            string dj_eja;
            string ld_name;
            string ld_jumin;
            string ld_hcode;
            string ld_addr;
            string ld_fymd;
            string ld_tymd;
            double ws_hm;
            double ws_mamt;
            double ws_gamt;
            double dj_tamt;
            double dj_wamt;
            double dj_eamt;
            double dj_gamt;
            double ld_hm;
            double ld_bamt;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value;
                ws_name = oForm.Items.Item("ws_name").Specific.Value;
                ws_jumin = oForm.Items.Item("ws_jumin").Specific.Value;
                ws_hcode = oForm.Items.Item("ws_hcode").Specific.Value;
                ws_hm = Convert.ToDouble(oForm.Items.Item("ws_hm").Specific.Value);
                ws_addr = oForm.Items.Item("ws_addr").Specific.Value;
                ws_fymd = oForm.Items.Item("ws_fymd").Specific.Value;
                ws_tymd = oForm.Items.Item("ws_tymd").Specific.Value;
                ws_mamt = Convert.ToDouble(oForm.Items.Item("ws_mamt").Specific.Value);
                ws_gamt = Convert.ToDouble(oForm.Items.Item("ws_gamt").Specific.Value);
                dj_name = oForm.Items.Item("dj_name").Specific.Value;
                dj_jumin = oForm.Items.Item("dj_jumin").Specific.Value;
                dj_fymd = oForm.Items.Item("dj_fymd").Specific.Value;
                dj_tymd = oForm.Items.Item("dj_tymd").Specific.Value;
                dj_eja = oForm.Items.Item("dj_eja").Specific.Value;
                dj_tamt = Convert.ToDouble(oForm.Items.Item("dj_tamt").Specific.Value);
                dj_wamt = Convert.ToDouble(oForm.Items.Item("dj_wamt").Specific.Value);
                dj_eamt = Convert.ToDouble(oForm.Items.Item("dj_eamt").Specific.Value);
                dj_gamt = Convert.ToDouble(oForm.Items.Item("dj_gamt").Specific.Value);
                ld_name = oForm.Items.Item("ld_name").Specific.Value;
                ld_jumin = oForm.Items.Item("ld_jumin").Specific.Value;
                ld_hcode = oForm.Items.Item("ld_hcode").Specific.Value;
                ld_hm = Convert.ToDouble(oForm.Items.Item("ld_hm").Specific.Value);
                ld_addr = oForm.Items.Item("ld_addr").Specific.Value;
                ld_fymd = oForm.Items.Item("ld_fymd").Specific.Value;
                ld_tymd = oForm.Items.Item("ld_tymd").Specific.Value;
                ld_bamt = Convert.ToDouble(oForm.Items.Item("ld_bamt").Specific.Value);
                seqn = oForm.Items.Item("LineNum").Specific.Value;

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

                if (string.IsNullOrEmpty(seqn))
                {
                    sQry = " Select case when max(isnull(seqn,0)) IS NULL OR max(isnull(seqn,0)) = 0 then 1 ";
                    sQry += "        else Convert(numeric(10), max(isnull(seqn, 0))) + 1 ";
                    sQry += "        end ";
                    sQry += "   From [p_seoyhouse] ";
                    sQry += "  Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    oRecordSet.DoQuery(sQry);
                    seqn = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                }

                sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And seqn = '" + seqn + "' And sabun = '" + sabun + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    sQry = "Update [p_seoyhouse] set ";
                    sQry += "ws_name = '" + ws_name + "',";
                    sQry += "ws_jumin = '" + ws_jumin + "',";
                    sQry += "ws_hcode = '" + ws_hcode + "',";
                    sQry += "ws_hm = '" + ws_hm + "',";
                    sQry += "ws_addr = '" + ws_addr + "',";
                    sQry += "ws_fymd = '" + ws_fymd + "',";
                    sQry += "ws_tymd = '" + ws_tymd + "',";
                    sQry += "ws_mamt = " + ws_mamt + ",";
                    sQry += "ws_gamt = " + ws_gamt + ",";
                    sQry += "dj_name = '" + dj_name + "',";
                    sQry += "dj_jumin = '" + dj_jumin + "',";
                    sQry += "dj_fymd = '" + dj_fymd + "',";
                    sQry += "dj_tymd = '" + dj_tymd + "',";
                    sQry += "dj_eja = '" + dj_eja + "',";
                    sQry += "dj_tamt = " + dj_tamt + ",";
                    sQry += "dj_wamt = " + dj_wamt + ",";
                    sQry += "dj_eamt = " + dj_eamt + ",";
                    sQry += "dj_gamt = " + dj_gamt + ",";
                    sQry += "ld_name = '" + ld_name + "',";
                    sQry += "ld_jumin = '" + ld_jumin + "',";
                    sQry += "ld_hcode = '" + ld_hcode + "',";
                    sQry += "ld_hm = '" + ld_hm + "',";
                    sQry += "ld_addr = '" + ld_addr + "',";
                    sQry += "ld_fymd = '" + ld_fymd + "',";
                    sQry += "ld_tymd = '" + ld_tymd + "',";
                    sQry += "ld_bamt = " + ld_bamt;
                    sQry += " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And seqn = '" + seqn + "' And sabun = '" + sabun + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY413_DataFind();
                }
                else
                {
                    sQry = "INSERT INTO [p_seoyhouse]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "seqn,";
                    sQry += "ws_name, ";
                    sQry += "ws_jumin, ";
                    sQry += "ws_hcode, ";
                    sQry += "ws_hm, ";
                    sQry += "ws_addr, ";
                    sQry += "ws_fymd, ";
                    sQry += "ws_tymd, ";
                    sQry += "ws_mamt, ";
                    sQry += "ws_gamt, ";
                    sQry += "dj_name, ";
                    sQry += "dj_jumin, ";
                    sQry += "dj_fymd, ";
                    sQry += "dj_tymd, ";
                    sQry += "dj_eja, ";
                    sQry += "dj_tamt, ";
                    sQry += "dj_wamt, ";
                    sQry += "dj_eamt, ";
                    sQry += "dj_gamt, ";
                    sQry += "ld_name, ";
                    sQry += "ld_jumin, ";
                    sQry += "ld_hcode, ";
                    sQry += "ld_hm, ";
                    sQry += "ld_addr, ";
                    sQry += "ld_fymd, ";
                    sQry += "ld_tymd, ";
                    sQry += "ld_bamt ";
                    sQry += " ) ";
                    sQry += "VALUES(";

                    sQry += "'" + saup + "',";
                    sQry += "'" + yyyy + "',";
                    sQry += "'" + sabun + "',";
                    sQry += "'" + seqn + "',";
                    sQry += "'" + ws_name + "',";
                    sQry += "'" + ws_jumin + "',";
                    sQry += "'" + ws_hcode + "',";
                    sQry += ws_hm + ",";
                    sQry += "'" + ws_addr + "',";
                    sQry += "'" + ws_fymd + "',";
                    sQry += "'" + ws_tymd + "',";
                    sQry += ws_mamt + ",";
                    sQry += ws_gamt + ",";

                    sQry += "'" + dj_name + "',";
                    sQry += "'" + dj_jumin + "',";
                    sQry += "'" + dj_fymd + "',";
                    sQry += "'" + dj_tymd + "',";
                    sQry += "'" + dj_eja + "',";
                    sQry += dj_tamt + ",";
                    sQry += dj_wamt + ",";
                    sQry += dj_eamt + ",";
                    sQry += dj_gamt + ",";

                    sQry += "'" + ld_name + "',";
                    sQry += "'" + ld_jumin + "',";
                    sQry += "'" + ld_hcode + "',";
                    sQry += ld_hm + ",";
                    sQry += "'" + ld_addr + "',";
                    sQry += "'" + ld_fymd + "',";
                    sQry += "'" + ld_tymd + "',";
                    sQry += ld_bamt;
                    sQry += " ) ";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY413_DataFind();
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
        /// PH_PY413_Delete 데이타 삭제
        /// </summary>
        private void PH_PY413_Delete()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (PSH_Globals.SBO_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    sQry = "Delete From [p_seoyhouse] Where saup = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                    sQry += "' And yyyy = '" + oForm.Items.Item("Year").Specific.Value.Trim();
                    sQry += "' And seqn = '" + oForm.Items.Item("LineNum").Specific.Value.Trim();
                    sQry += "' And sabun = '" + oForm.Items.Item("MSTCOD").Specific.Value.Trim() + "'";
                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY413_DataFind();
                    oForm.ActiveItem = "MSTCOD";
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
                        PH_PY413_DataFind();
                    }
                    if (pVal.ItemUID == "Btn01")  // 저장
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value.Trim() + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("등록불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY413_SAVE();
                        }
                    }
                    if (pVal.ItemUID == "Btn_del")  // 삭제
                    {
                        sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + oForm.Items.Item("Year").Specific.Value.Trim() + "'";
                        oRecordSet.DoQuery(sQry);

                        Result = oRecordSet.Fields.Item(0).Value;
                        if (Result != "Y")
                        {
                            PSH_Globals.SBO_Application.MessageBox("삭제불가 년도입니다. 담당자에게 문의바랍니다.");
                        }
                        if (Result == "Y")
                        {
                            PH_PY413_Delete();
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
            string yyyy;
            string CLTCOD;
            string MSTCOD;
            string FullName;
            double amt;
            double gamt;
            double samt;
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

                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "FullName":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();

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
                                sQry += " And U_status <> '5'";    // 퇴사자 제외
                                sQry += " and U_FullName = '" + FullName + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                //                            oForm.Items("MSTCOD").Specific.Value = oRecordSet.Fields("Code").Value
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "ws_mamt":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();

                                //총급여액계산해서 5500 이하는 12% 아니면 10%
                                sQry = "SELECT SUM(gwase) ";
                                sQry += "FROM( SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.Code ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "      Union All ";
                                sQry += "      SELECT gwase   = SUM( a.U_GWASEE ) ";
                                sQry += "        FROM [@PH_PY112A] a Inner Join [@PH_PY001A] b On a.U_MSTCOD = b.U_PreCode ";
                                sQry += "       WHERE b.U_CLTCOD = '" + CLTCOD + "' ";
                                sQry += "         And a.U_CLTCOD = b.U_CLTCOD ";
                                sQry += "         And a.U_YM     BETWEEN  '" + yyyy + "' + '01' AND '" + yyyy + "' + '12' ";
                                sQry += "         And isnull(b.Code,'')  = '" + MSTCOD + "' ";
                                sQry += "         And Isnull(b.U_PreCode,'') <> '' ";
                                sQry += "     ) g";

                                oRecordSet.DoQuery(sQry);
                                samt = oRecordSet.Fields.Item(0).Value;  // 총급여액(과세대상)

                                if (Convert.ToDouble(oForm.Items.Item("ws_mamt").Specific.Value.Trim()) > 7500000)  // 한도 7백5십만원
                                {
                                    oForm.Items.Item("ws_mamt").Specific.Value = 7500000;
                                }

                                gamt = 0;
                                
                                if (samt > 70000000)  // 7천초과자 0
                                {
                                    gamt = 0;
                                }
                                else if (samt <= 70000000)  // 7천이하자  12%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt").Specific.Value.Trim()) * 0.12, 0);
                                }
                                else if (samt <= 55000000)  // 5천5백이하자  15%
                                {
                                    gamt = System.Math.Round(Convert.ToDouble(oForm.Items.Item("ws_mamt").Specific.Value.Trim()) * 0.15, 0);
                                }

                                if (gamt < 0)
                                {
                                    gamt = 0;
                                }

                                oForm.Items.Item("ws_gamt").Specific.Value = gamt;
                                break;

                            case "dj_tamt":
                                amt = 0;
                                gamt = 0;
                                amt = Convert.ToDouble(oForm.Items.Item("dj_tamt").Specific.Value.Trim());
                                gamt = System.Math.Round(amt * 0.4, 0);
                                if (gamt > 3000000)
                                {
                                    oForm.Items.Item("dj_gamt").Specific.Value = 3000000;
                                }
                                else
                                {
                                    oForm.Items.Item("dj_gamt").Specific.Value = gamt;
                                }
                                break;
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
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row >= 0)
                        {
                            oForm.Freeze(true);

                            Param01 = oDS_PH_PY413.Columns.Item("사업장").Cells.Item(pVal.Row).Value;
                            Param02 = oDS_PH_PY413.Columns.Item("년도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY413.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY413.Columns.Item("순번").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY413_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.RecordCount == 0)
                            {
                                oForm.DataSources.UserDataSources.Item("ws_name").Value = "";
                                oForm.DataSources.UserDataSources.Item("ws_jumin").Value = "";
                                oForm.Items.Item("ws_hcode").Specific.Select(oRecordSet.Fields.Item("ws_hcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("ws_hm").Value = "";
                                oForm.DataSources.UserDataSources.Item("ws_addr").Value ="";
                                oForm.DataSources.UserDataSources.Item("ws_fymd").Value ="";
                                oForm.DataSources.UserDataSources.Item("ws_tymd").Value ="";
                                oForm.DataSources.UserDataSources.Item("ws_mamt").Value ="";
                                oForm.DataSources.UserDataSources.Item("ws_gamt").Value ="";
                                oForm.DataSources.UserDataSources.Item("dj_name").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_jumin").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_fymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_tymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_eja").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_tamt").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_wamt").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_eamt").Value = "";
                                oForm.DataSources.UserDataSources.Item("dj_gamt").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_name").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_jumin").Value = "";
                                oForm.Items.Item("ld_hcode").Specific.Select(oRecordSet.Fields.Item("ld_hcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("ld_hm").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_addr").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_fymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_tymd").Value = "";
                                oForm.DataSources.UserDataSources.Item("ld_bamt").Value = "";
                                oForm.DataSources.UserDataSources.Item("LineNum").Value = "";
                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }
                            else
                            {
                                sQry = "  Select Code,";
                                sQry += "        FullName = U_FullName,";
                                sQry += "        TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                             WHERE Code = '1'";
                                sQry += "                             And U_Code = U_TeamCode),''),";
                                sQry += "        RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                            From [@PS_HR200L]";
                                sQry += "                            WHERE Code = '2'";
                                sQry += "                            And U_Code = U_RspCode),''),";
                                sQry += "        ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                             WHERE Code = '9'";
                                sQry += "                             And U_Code  = U_ClsCode";
                                sQry += "                             And U_Char3 = U_CLTCOD),'')";
                                sQry += " From [@PH_PY001A]";
                                sQry += " Where U_CLTCOD = '" + Param01 + "'";
                                sQry += "  and  Code = '" + Param03 + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.Items.Item("MSTCOD").Specific.Value = oRecordSet.Fields.Item("Code").Value;
                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;

                                sQry = " Select * From [p_seoyhouse] Where saup = '" + Param01 + "' And yyyy = '" + Param02 + "' And sabun = '" + Param03 + "' And seqn = '" + Param04 + "'";
                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("ws_name").Value = oRecordSet.Fields.Item("ws_name").Value;
                                oForm.DataSources.UserDataSources.Item("ws_jumin").Value = oRecordSet.Fields.Item("ws_jumin").Value;
                                oForm.Items.Item("ws_hcode").Specific.Select(oRecordSet.Fields.Item("ws_hcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("ws_hm").Value = oRecordSet.Fields.Item("ws_hm").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ws_addr").Value = oRecordSet.Fields.Item("ws_addr").Value;
                                oForm.DataSources.UserDataSources.Item("ws_fymd").Value = oRecordSet.Fields.Item("ws_fymd").Value;
                                oForm.DataSources.UserDataSources.Item("ws_tymd").Value = oRecordSet.Fields.Item("ws_tymd").Value;
                                oForm.DataSources.UserDataSources.Item("ws_mamt").Value = oRecordSet.Fields.Item("ws_mamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ws_gamt").Value = oRecordSet.Fields.Item("ws_gamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("dj_name").Value = oRecordSet.Fields.Item("dj_name").Value;
                                oForm.DataSources.UserDataSources.Item("dj_jumin").Value = oRecordSet.Fields.Item("dj_jumin").Value;
                                oForm.DataSources.UserDataSources.Item("dj_fymd").Value = oRecordSet.Fields.Item("dj_fymd").Value;
                                oForm.DataSources.UserDataSources.Item("dj_tymd").Value = oRecordSet.Fields.Item("dj_tymd").Value;
                                oForm.DataSources.UserDataSources.Item("dj_eja").Value = oRecordSet.Fields.Item("dj_eja").Value;
                                oForm.DataSources.UserDataSources.Item("dj_tamt").Value = oRecordSet.Fields.Item("dj_tamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("dj_wamt").Value = oRecordSet.Fields.Item("dj_wamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("dj_eamt").Value = oRecordSet.Fields.Item("dj_eamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("dj_gamt").Value = oRecordSet.Fields.Item("dj_gamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ld_name").Value = oRecordSet.Fields.Item("ld_name").Value;
                                oForm.DataSources.UserDataSources.Item("ld_jumin").Value = oRecordSet.Fields.Item("ld_jumin").Value;
                                oForm.Items.Item("ld_hcode").Specific.Select(oRecordSet.Fields.Item("ld_hcode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.DataSources.UserDataSources.Item("ld_hm").Value = oRecordSet.Fields.Item("ld_hm").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("ld_addr").Value = oRecordSet.Fields.Item("ld_addr").Value;
                                oForm.DataSources.UserDataSources.Item("ld_fymd").Value = oRecordSet.Fields.Item("ld_fymd").Value;
                                oForm.DataSources.UserDataSources.Item("ld_tymd").Value = oRecordSet.Fields.Item("ld_tymd").Value;
                                oForm.DataSources.UserDataSources.Item("ld_bamt").Value = oRecordSet.Fields.Item("ld_bamt").Value.ToString();
                                oForm.DataSources.UserDataSources.Item("LineNum").Value = oRecordSet.Fields.Item("Seqn").Value.ToString();

                                oForm.Items.Item("CLTCOD").Enabled = false;
                                oForm.Items.Item("Year").Enabled = false;
                                oForm.Items.Item("MSTCOD").Enabled = false;
                            }
                            oForm.ActiveItem = "ws_name";
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY413);
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
                            PH_PY413_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY413_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
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
    }
}
