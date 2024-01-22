using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정산기부금등록
    /// </summary>
    internal class PH_PY407 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DataTable oDS_PH_PY407A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY407L;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY407.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY407_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY407");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY407_CreateItems();
                PH_PY407_FormItemEnabled();
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
        private void PH_PY407_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PH_PY407L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oGrid1 = oForm.Items.Item("Grid01").Specific;

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oForm.DataSources.DataTables.Add("PH_PY407");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY407");
                oDS_PH_PY407A = oForm.DataSources.DataTables.Item("PH_PY407");

                // 그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("연도", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("관계", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("관계명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("주민번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부내용", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사업자번호", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부처명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금액(국세청)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부금액(국세청외)", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("기부장려금신청금액", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY407").Columns.Add("사업장", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                
                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //성명
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
                oForm.Items.Item("rel").Specific.ValidValues.Add("", "");
                oForm.Items.Item("rel").Specific.ValidValues.Add("1", "거주자");
                oForm.Items.Item("rel").Specific.ValidValues.Add("2", "배우자");
                oForm.Items.Item("rel").Specific.ValidValues.Add("3", "직계비속");
                oForm.Items.Item("rel").Specific.ValidValues.Add("4", "직계존속");
                oForm.Items.Item("rel").Specific.ValidValues.Add("5", "형제,자매");
                oForm.Items.Item("rel").Specific.ValidValues.Add("6", "그외");
                oForm.Items.Item("rel").DisplayDesc = true;
                oForm.Items.Item("rel").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                
                // 성명
                oForm.DataSources.UserDataSources.Add("kname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("kname").Specific.DataBind.SetBound(true, "", "kname");

                // 주민번호
                oForm.DataSources.UserDataSources.Add("juminno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("juminno").Specific.DataBind.SetBound(true, "", "juminno");

                // 기부금코드  73
                oForm.DataSources.UserDataSources.Add("gibucd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y' Order by U_Num1 ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("gibucd").Specific, "Y");

                // 기부내용 2018추가
                oForm.DataSources.UserDataSources.Add("gibudscr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("gibudscr").Specific.ValidValues.Add("1", "금전");
                oForm.Items.Item("gibudscr").Specific.ValidValues.Add("2", "현물");
                oForm.Items.Item("gibudscr").DisplayDesc = true;
                oForm.Items.Item("gibudscr").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 사업자번호
                oForm.DataSources.UserDataSources.Add("saupno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 13);
                oForm.Items.Item("saupno").Specific.DataBind.SetBound(true, "", "saupno");

                // 기부처명
                oForm.DataSources.UserDataSources.Add("sangho", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 60);
                oForm.Items.Item("sangho").Specific.DataBind.SetBound(true, "", "sangho");

                // 공제금액(국세청)
                oForm.DataSources.UserDataSources.Add("ntamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("ntamt").Specific.DataBind.SetBound(true, "", "ntamt");

                // 공제금액(국세청외)
                oForm.DataSources.UserDataSources.Add("amt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("amt").Specific.DataBind.SetBound(true, "", "amt");

                // 기부장려금신청금액 2016
                oForm.DataSources.UserDataSources.Add("jamt", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("jamt").Specific.DataBind.SetBound(true, "", "jamt");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY407_FormItemEnabled()
        {
            try
            {
                oForm.EnableMenu("1282", true);      // 문서추가

                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("MSTCOD").Specific.Value = "";
                    oForm.Items.Item("FullName").Specific.Value = "";
                    oForm.Items.Item("TeamName").Specific.Value = "";
                    oForm.Items.Item("RspName").Specific.Value = "";
                    oForm.Items.Item("ClsName").Specific.Value = "";
                }

                oForm.Items.Item("rel").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("gibudscr").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                oForm.DataSources.UserDataSources.Item("sangho").Value = "";

                oForm.DataSources.UserDataSources.Item("ntamt").Value = "0";
                oForm.DataSources.UserDataSources.Item("amt").Value = "0";
                oForm.DataSources.UserDataSources.Item("jamt").Value = "0";

                oForm.Items.Item("CLTCOD").Enabled = true;
                oForm.Items.Item("Year").Enabled = true;
                oForm.Items.Item("MSTCOD").Enabled = true;

                oForm.Items.Item("juminno").Enabled = true;
                oForm.Items.Item("saupno").Enabled = true;
                oForm.Items.Item("gibucd").Enabled = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY407_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY407_DataFind
        /// </summary>
        private void PH_PY407_DataFind()
        {
            string sQry;

            try
            {
                oForm.Freeze(true);
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장이 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return;
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도가 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return;
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번이 없습니다. 확인바랍니다..", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return;
                }

                PH_PY407_FormItemEnabled();

                sQry = "EXEC PH_PY407_01 '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "', '" + oForm.Items.Item("MSTCOD").Specific.Value + "'";
                oDS_PH_PY407A.ExecuteQuery(sQry);
                PH_PY407_TitleSetting();
                oGrid1.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_SAVE
        /// </summary>
        private void PH_PY407_SAVE()
        {
            // 데이타 저장
            string sQry;
            string saup;
            string yyyy;
            string sabun;
            string rel;
            string kname;
            string juminno;
            string gibucd;
            string gibudscr;
            string saupno;
            string sangho;
            string FullName;
            double Amt;
            double ntamt;
            double jamt;
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
                gibucd = oForm.Items.Item("gibucd").Specific.Value.ToString().Trim();
                gibudscr = oForm.Items.Item("gibudscr").Specific.Value.ToString().Trim();
                saupno = oForm.Items.Item("saupno").Specific.Value.ToString().Trim();
                sangho = oForm.Items.Item("sangho").Specific.Value.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();
                ntamt = Convert.ToDouble(oForm.Items.Item("ntamt").Specific.Value);
                Amt = Convert.ToDouble(oForm.Items.Item("amt").Specific.Value);
                jamt = Convert.ToDouble(oForm.Items.Item("jamt").Specific.Value);

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
                if (string.IsNullOrEmpty(juminno) || (ntamt == 0 && Amt == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("정상적인 자료가 아닙니다. 확인바랍니다..");
                    return;
                }

                sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                sQry += " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    // 갱신
                    sQry = "Update [p_seoygibuhis] set ";
                    sQry += "rel = '" + rel + "',";
                    sQry += "kname = '" + kname + "',";
                    sQry += "sangho = '" + sangho + "',";
                    sQry += "gibudscr = '" + gibudscr + "',";
                    sQry += "ntamt = " + ntamt + ",";
                    sQry += "jamt = " + jamt + ",";
                    sQry += "amt =" + Amt;
                    sQry += " Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                    sQry += " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    PH_PY407_DataFind();
                }
                else
                {
                    // 신규
                    sQry = "INSERT INTO [p_seoygibuhis]";
                    sQry += " (";
                    sQry += "saup,";
                    sQry += "yyyy,";
                    sQry += "sabun,";
                    sQry += "rel,";
                    sQry += "kname,";
                    sQry += "juminno,";
                    sQry += "gibucd,";
                    sQry += "gibudscr,";
                    sQry += "saupno,";
                    sQry += "sangho,";
                    sQry += "ntamt,";
                    sQry += "jamt,";
                    sQry += "amt)";
                    sQry += " VALUES(";
                    sQry += "'" + saup + "',";
                    sQry += "'" + yyyy + "',";
                    sQry += "'" + sabun + "',";
                    sQry += "'" + rel + "',";
                    sQry += "'" + kname + "',";
                    sQry += "'" + juminno + "',";
                    sQry += "'" + gibucd + "',";
                    sQry += "'" + gibudscr + "',";
                    sQry += "'" + saupno + "',";
                    sQry += "'" + sangho + "',";
                    sQry += ntamt + ",";
                    sQry += jamt + ",";
                    sQry += Amt + " )";

                    oRecordSet.DoQuery(sQry);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY407_DataFind();
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_Delete
        /// </summary>
        private void PH_PY407_Delete()
        {
            // 데이타 삭제
            string sQry;
            string saup;
            string yyyy;
            string sabun;
            string gibucd;
            string saupno;
            string juminno;
            string FullName;
            double cnt;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                saup = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                yyyy = oForm.Items.Item("Year").Specific.Value.ToString().Trim();
                sabun = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                juminno = oForm.Items.Item("juminno").Specific.Value.ToString().Trim();
                gibucd = oForm.Items.Item("gibucd").Specific.Value.ToString().Trim();
                saupno = oForm.Items.Item("saupno").Specific.Value.ToString().Trim();
                FullName = oForm.Items.Item("kname").Specific.Value.ToString().Trim();

                sQry = " Select Count(*) From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                sQry += " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";

                oRecordSet.DoQuery(sQry);

                cnt = oRecordSet.Fields.Item(0).Value;
                if (cnt > 0)
                {

                    if (PSH_Globals.SBO_Application.MessageBox(" 선택한대상자('" + FullName + "')을 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                    {
                        sQry = "Delete From [p_seoygibuhis] Where saup = '" + saup + "' And yyyy = '" + yyyy + "' And sabun = '" + sabun + "'";
                        sQry += " And saupno = '" + saupno + "' And gibucd = '" + gibucd + "' And  juminno = '" + juminno + "'";
                        oRecordSet.DoQuery(sQry);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PH_PY407_DataFind();
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY407_TitleSetting
        /// </summary>
        private void PH_PY407_TitleSetting()
        {
            int i;
            string[] COLNAM = new string[15];

            try
            {
                COLNAM[0] = "연도";
                COLNAM[1] = "관계";
                COLNAM[2] = "관계명";
                COLNAM[3] = "성명";
                COLNAM[4] = "주민번호";
                COLNAM[5] = "기부금코드";
                COLNAM[6] = "기부금명";
                COLNAM[7] = "기부내용";
                COLNAM[8] = "사업자번호";
                COLNAM[9] = "기부처명";
                COLNAM[10] = "기부금액(국세청)";
                COLNAM[11] = "기부금액(국세청외)";
                COLNAM[12] = "기부장려금신청금액";
                COLNAM[13] = "사번";
                COLNAM[14] = "사업장";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    oGrid1.Columns.Item(i).Editable = false;
                }
                oGrid1.Columns.Item(10).RightJustified = true;
                oGrid1.Columns.Item(11).RightJustified = true;
                oGrid1.Columns.Item(12).RightJustified = true;
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY407_TitleSetting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        PH_PY407_DataFind();
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
                            PH_PY407_SAVE();
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
                            PH_PY407_Delete();
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
            string MSTCOD;
            string relate;
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
                            oDS_PH_PY407L.Clear();

                            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;
                            relate = oForm.Items.Item("rel").Specific.Value;

                            sQry = "EXEC [PH_PY407_03] '" + MSTCOD + "', '" + relate + "'";

                            oRecordSet.DoQuery(sQry);

                            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                            {
                                if (i + 1 > oDS_PH_PY407L.Size)
                                {
                                    oDS_PH_PY407L.InsertRecord(i);
                                }

                                oMat01.AddRow();
                                oDS_PH_PY407L.Offset = i;
                                oDS_PH_PY407L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                oDS_PH_PY407L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("kname").Value.ToString().Trim());
                                oDS_PH_PY407L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("juminno").Value.ToString().Trim());
                                oDS_PH_PY407L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("birthymd").Value.ToString().Trim());
                                oDS_PH_PY407L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("relate").Value.ToString().Trim());
                                oDS_PH_PY407L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("addr").Value.ToString().Trim());
                                oRecordSet.MoveNext();
                            }

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();

                            if (oRecordSet.RecordCount == 0)
                            {
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                                oForm.DataSources.UserDataSources.Item("sangho").Value = "";

                                oForm.DataSources.UserDataSources.Item("ntamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("amt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("jamt").Value = "0";

                            }

                            if (oRecordSet.RecordCount == 1)
                            {
                                oForm.Items.Item("kname").Specific.Value = oMat01.Columns.Item("kname").Cells.Item(1).Specific.Value;
                                oForm.Items.Item("juminno").Specific.Value = oMat01.Columns.Item("juminno").Cells.Item(1).Specific.Value;
                            }
                        }

                        if (pVal.ItemUID == "gibucd")
                        {
                            if ((oForm.Items.Item("gibucd").Specific.Value.ToString().Trim() == "20" ||  // 정치자금기부금
                                     oForm.Items.Item("gibucd").Specific.Value.ToString().Trim() == "42" ||  // 우리사주조합기부금
                                        oForm.Items.Item("gibucd").Specific.Value.ToString().Trim() == "43") && // 고향사랑기부금
                                           oForm.Items.Item("rel").Specific.Value.ToString().Trim() != "1")   // 관계가 본인이 아닐시
                            {
                                PSH_Globals.SBO_Application.MessageBox("정치자금, 우리사주조합, 고향사랑 기부금은 본인(거주자)만 가능 합니다. 확인하세요.");
                                oForm.DataSources.UserDataSources.Item("gibucd").Value = "";
                                oForm.Items.Item("gibucd").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
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
                                sQry = "  Select    Code,";
                                sQry += "           FullName = U_FullName,";
                                sQry += "           TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '1'";
                                sQry += "                                   And U_Code = U_TeamCode),''),";
                                sQry += "           RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '2'";
                                sQry += "                                   And U_Code = U_RspCode),''),";
                                sQry += "           ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '9'";
                                sQry += "                                   And U_Code  = U_ClsCode";
                                sQry += "                                   And U_Char3 = U_CLTCOD),'')";
                                sQry += " From      [@PH_PY001A]";
                                sQry += " Where     U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += "           and Code = '" + oForm.Items.Item("MSTCOD").Specific.Value + "'";

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "FullName":
                                sQry = "  Select    Code,";
                                sQry += "           FullName = U_FullName,";
                                sQry += "           TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '1'";
                                sQry += "                                   And U_Code = U_TeamCode),''),";
                                sQry += "           RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '2'";
                                sQry += "                                   And U_Code = U_RspCode),''),";
                                sQry += "           ClsName  = Isnull((SELECT U_CodeNm";
                                sQry += "                               From [@PS_HR200L]";
                                sQry += "                               WHERE Code = '9'";
                                sQry += "                                   And U_Code  = U_ClsCode";
                                sQry += "                                   And U_Char3 = U_CLTCOD),'')";
                                sQry += " From      [@PH_PY001A]";
                                sQry += " Where     U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                                sQry += "               And U_status <> '5'";
                                sQry += "               and U_FullName = '" + oForm.Items.Item("FullName").Specific.Value + "'"; //퇴사자 제외

                                oRecordSet.DoQuery(sQry);

                                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("Code").Value;
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value;
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value;
                                oForm.Items.Item("ClsName").Specific.Value = oRecordSet.Fields.Item("ClsName").Value;
                                break;

                            case "ntamt":
                                break;
                            case "juminno":
                                //주민번호
                                //주민번호입력시 생년월일 생성
                                if (oForm.Items.Item("juminno").Specific.Value.ToString().Trim().Length != 13)
                                {
                                    PSH_Globals.SBO_Application.MessageBox("주민번호자릿수가 틀립니다. 확인하세요.");
                                }
                                else
                                {
                                }
                                break;
                            case "amt":
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
            string sQry; 
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
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

                            Param01 = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                            Param02 = oDS_PH_PY407A.Columns.Item("연도").Cells.Item(pVal.Row).Value;
                            Param03 = oDS_PH_PY407A.Columns.Item("사번").Cells.Item(pVal.Row).Value;
                            Param04 = oDS_PH_PY407A.Columns.Item("사업자번호").Cells.Item(pVal.Row).Value;
                            Param05 = oDS_PH_PY407A.Columns.Item("기부금코드").Cells.Item(pVal.Row).Value;
                            Param06 = oDS_PH_PY407A.Columns.Item("주민번호").Cells.Item(pVal.Row).Value;

                            sQry = "EXEC PH_PY407_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "'";
                            oRecordSet.DoQuery(sQry);

                            if (oRecordSet.RecordCount == 0)
                            {
                                oForm.Items.Item("rel").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("kname").Value = "";
                                oForm.DataSources.UserDataSources.Item("juminno").Value = "";
                                oForm.Items.Item("gibucd").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("gibudscr").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.DataSources.UserDataSources.Item("saupno").Value = "";
                                oForm.DataSources.UserDataSources.Item("sangho").Value = "";
                                oForm.DataSources.UserDataSources.Item("ntamt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("amt").Value = "0";
                                oForm.DataSources.UserDataSources.Item("jamt").Value = "0";

                                PSH_Globals.SBO_Application.MessageBox("결과가 존재하지 않습니다.");
                            }

                            oForm.Items.Item("gibucd").Specific.Select(oRecordSet.Fields.Item("gibucd").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("gibudscr").Specific.Select(oRecordSet.Fields.Item("gibudscr").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("rel").Specific.Select(oRecordSet.Fields.Item("rel").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.DataSources.UserDataSources.Item("kname").Value = oRecordSet.Fields.Item("kname").Value;
                            oForm.DataSources.UserDataSources.Item("juminno").Value = oRecordSet.Fields.Item("juminno").Value;
                            oForm.DataSources.UserDataSources.Item("saupno").Value = oRecordSet.Fields.Item("saupno").Value;
                            oForm.DataSources.UserDataSources.Item("sangho").Value = oRecordSet.Fields.Item("sangho").Value;
                            oForm.DataSources.UserDataSources.Item("ntamt").Value = oRecordSet.Fields.Item("ntamt").Value.ToString();
                            oForm.DataSources.UserDataSources.Item("amt").Value = oRecordSet.Fields.Item("amt").Value.ToString();
                            oForm.DataSources.UserDataSources.Item("jamt").Value = oRecordSet.Fields.Item("jamt").Value.ToString();

                            oForm.ActiveItem = "rel";

                            oForm.Items.Item("CLTCOD").Enabled = false;
                            oForm.Items.Item("Year").Enabled = false;
                            oForm.Items.Item("MSTCOD").Enabled = false;

                            oForm.Items.Item("gibucd").Enabled = false;
                            oForm.Items.Item("juminno").Enabled = false;
                            oForm.Items.Item("saupno").Enabled = false;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY407L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY407A);
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
                            PH_PY407_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY407_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY407_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY407_FormItemEnabled();
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
