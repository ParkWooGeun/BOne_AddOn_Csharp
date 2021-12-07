using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 사용자 권한 등록
    /// </summary>
    internal class PS_SY015 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PS_SY015A;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        private List<string> temp = new List<string>();

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SY015.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_SY015_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_SY015");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_SY015_CreateItems();
                PS_SY015_FormItemEnabled();
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
        private void PS_SY015_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_SY015");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PS_SY015");
                oDS_PS_SY015A = oForm.DataSources.DataTables.Item("PS_SY015");
                
                // 구분
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("B", "기본");
                oForm.Items.Item("pGubun").Specific.ValidValues.Add("H", "인사");
                oForm.Items.Item("pGubun").DisplayDesc = true;

                // 폴더/화면구분
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("F", "폴더");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("S", "화면");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("C", "복제");
                oForm.Items.Item("pFSGubun").Specific.ValidValues.Add("A", "권한");
                oForm.Items.Item("pFSGubun").DisplayDesc = true;

                // 사용여부
                oForm.Items.Item("UseYN").Specific.ValidValues.Add("Y", "사용");
                oForm.Items.Item("UseYN").Specific.ValidValues.Add("N", "미사용");
                oForm.Items.Item("UseYN").DisplayDesc = true;

                //변경타입
                oForm.Items.Item("MType").Specific.ValidValues.Add("N", "신규");
                oForm.Items.Item("MType").Specific.ValidValues.Add("M", "변경");
                oForm.Items.Item("MType").Specific.ValidValues.Add("C", "부서이동");
                oForm.Items.Item("MType").DisplayDesc = true;

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A006'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Modual").Specific,  "Y");
                oForm.Items.Item("Modual").DisplayDesc = true;

                // Position
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='A005'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Position").Specific, "Y");
                oForm.Items.Item("Position").DisplayDesc = true;

                // Sub1
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub1").Specific, "Y");
                oForm.Items.Item("Sub1").DisplayDesc = true;

                // Sub2
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub2").Specific, "Y");
                oForm.Items.Item("Sub2").DisplayDesc = true;

                // Sub3
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Sub3").Specific, "Y");
                oForm.Items.Item("Sub3").DisplayDesc = true;

                // 순서
                sQry = "select U_Minor , U_CdName from [@PS_SY001L] where code ='H002'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("No").Specific, "Y");
                oForm.Items.Item("No").DisplayDesc = true;

                // Level
                oForm.Items.Item("Level").Specific.ValidValues.Add("0", "0");
                oForm.Items.Item("Level").Specific.ValidValues.Add("1", "1");
                oForm.Items.Item("Level").Specific.ValidValues.Add("2", "2");
                oForm.Items.Item("Level").DisplayDesc = true;

                // FatherId
                sQry = "select  distinct t.a,t.b   from (select distinct UniqueID as a , UniqueID as b from Authority_Folder union all select distinct FatherID as a , FatherID as b from Authority_Folder) t";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FatherID").Specific, "Y");
                oForm.Items.Item("FatherID").DisplayDesc = true;

                // String
                oForm.DataSources.UserDataSources.Add("Strings", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("Strings").Specific.DataBind.SetBound(true, "", "Strings");

                // UniqueID
                oForm.DataSources.UserDataSources.Add("UniqueID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("UniqueID").Specific.DataBind.SetBound(true, "", "UniqueID");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PS_SY015_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Items.Item("pFSGubun").Specific.Value == "F")
                {
                    oForm.Items.Item("pUserID").Enabled = false;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = true;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;
                    oForm.Items.Item("Btn_Find").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;
                    oForm.Items.Item("pUnique").Enabled = false;
                    oForm.Items.Item("refData").Enabled = false;
                    oForm.Items.Item("UseYN").Enabled = true;
                    oForm.Items.Item("Btn01").Enabled = true;
                }
                else if (oForm.Items.Item("pFSGubun").Specific.Value == "S")
                {
                    oForm.Items.Item("pUserID").Enabled = false;
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = true;
                    oForm.Items.Item("Sub1").Enabled = true;
                    oForm.Items.Item("Sub2").Enabled = true;
                    oForm.Items.Item("Sub3").Enabled = true;
                    oForm.Items.Item("No").Enabled = true;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = true;
                    oForm.Items.Item("FatherID").Enabled = true;
                    oForm.Items.Item("Position").Enabled = true;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = true;
                    oForm.Items.Item("Btn_Find").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;
                    oForm.Items.Item("pUnique").Enabled = true;
                    oForm.Items.Item("refData").Enabled = true;
                    oForm.Items.Item("UseYN").Enabled = true;
                    oForm.Items.Item("Btn01").Enabled = true;
                }
                else if (oForm.Items.Item("pFSGubun").Specific.Value == "C")
                {
                    oForm.Items.Item("CPUserID").Enabled = true;
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("pGubun").Enabled = false;
                    oForm.Items.Item("Modual").Enabled = false;
                    oForm.Items.Item("Sub1").Enabled = false;
                    oForm.Items.Item("Sub2").Enabled = false;
                    oForm.Items.Item("Sub3").Enabled = false;
                    oForm.Items.Item("No").Enabled = false;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = false;
                    oForm.Items.Item("FatherID").Enabled = false;
                    oForm.Items.Item("Position").Enabled = false;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;
                    oForm.Items.Item("Btn_Find").Enabled = false;
                    oForm.Items.Item("Bt_Copy").Enabled = true;
                    oForm.Items.Item("pUnique").Enabled = false;
                    oForm.Items.Item("refData").Enabled = true;
                    oForm.Items.Item("UseYN").Enabled = false;
                    oForm.Items.Item("Btn01").Enabled = false;
                }
                else if (oForm.Items.Item("pFSGubun").Specific.Value == "A")
                {
                    oForm.Items.Item("CPUserID").Enabled = false;
                    oForm.Items.Item("pUserID").Enabled = true;
                    oForm.Items.Item("pGubun").Enabled = true;
                    oForm.Items.Item("Modual").Enabled = false;
                    oForm.Items.Item("Sub1").Enabled = false;
                    oForm.Items.Item("Sub2").Enabled = false;
                    oForm.Items.Item("Sub3").Enabled = false;
                    oForm.Items.Item("No").Enabled = false;
                    oForm.Items.Item("Level").Enabled = false;
                    oForm.Items.Item("Strings").Enabled = false;
                    oForm.Items.Item("FatherID").Enabled = false;
                    oForm.Items.Item("Position").Enabled = false;
                    oForm.Items.Item("UserID").Enabled = false;
                    oForm.Items.Item("Sequence").Enabled = false;
                    oForm.Items.Item("UniqueID").Enabled = false;
                    oForm.Items.Item("Btn_Find").Enabled = true;
                    oForm.Items.Item("Bt_Copy").Enabled = false;
                    oForm.Items.Item("pUnique").Enabled = true;
                    oForm.Items.Item("refData").Enabled = true;
                    oForm.Items.Item("UseYN").Enabled = false;
                    oForm.Items.Item("Btn01").Enabled = true;
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", false); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", false); // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); // 문서찾기
                    oForm.EnableMenu("1282", true); // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPS_SY015_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_SY015_MTX01
        /// </summary>
        private void PS_SY015_MTX01()
        {
            int i;
            int iRow;
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("pGubun").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("pUserID").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("pUnique").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(Param01.ToString().Trim()))
                {
                    errMessage = "구분이 없습니다. 확인바랍니다.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim() == "A")
                {
                    if (string.IsNullOrEmpty(Param03.ToString().Trim()) && string.IsNullOrEmpty(Param04.ToString().Trim()))
                    {
                        errMessage = "UserID 또는 화면코드는 필수입니다.";
                        throw new Exception();
                    }
                }
                sQry = "EXEC PS_SY015_01 '" + Param01 + "', '" + Param02 + "', '"+  Param03 + "', '" + Param04 + "'";
                oDS_PS_SY015A.ExecuteQuery(sQry);
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PS_SY015_TitleSetting(iRow);
                if(oForm.Items.Item("pFSGubun").Specific.value == "A")
                {
                    temp.Clear();
                    for (i = 0; i < oGrid1.Rows.Count; i++)
                    {
                        temp.Add(oGrid1.DataTable.Columns.Item("UniqueID").Cells.Item(i).Value + "_" + oGrid1.DataTable.Columns.Item("UserID").Cells.Item(i).Value + "_" + oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value);
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_SY015_MTX02
        /// </summary>
        private void PS_SY015_MTX02(string oUID, int oRow, string oCol)
        {
            int sRow;
            string sQry;
            string Param01;
            string Param02;
            string Param03 = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sRow = oRow;
                Param01 = oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                Param02 = oDS_PS_SY015A.Columns.Item("UniqueID").Cells.Item(oRow).Value;
                if (Param01 == "A")
                {
                    Param03 = oDS_PS_SY015A.Columns.Item("UserID").Cells.Item(oRow).Value;
                }
                sQry = "EXEC PS_SY015_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                // Screen일때 UserID를 가져옴.
                if (Param01 == "A")
                {
                    oForm.Items.Item("UserID").Specific.value = oRecordSet.Fields.Item("UserID").Value;
                }
                // Folder일때 Level을 가져옴.
                else if (Param01 == "F")
                {
                    oForm.Items.Item("Level").Specific.Select(oRecordSet.Fields.Item("Level").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                //공통 S
                oForm.Items.Item("Modual").Specific.Select(oRecordSet.Fields.Item("Modual").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub1").Specific.Select(oRecordSet.Fields.Item("Sub1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub2").Specific.Select(oRecordSet.Fields.Item("Sub2").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Sub3").Specific.Select(oRecordSet.Fields.Item("Sub3").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("No").Specific.Select(oRecordSet.Fields.Item("No").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("Position").Specific.Select(oRecordSet.Fields.Item("Position").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("FatherID").Specific.Select(oRecordSet.Fields.Item("FatherID").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("Strings").Value = oRecordSet.Fields.Item("Strings").Value;
                oForm.DataSources.UserDataSources.Item("UniqueId").Value = oRecordSet.Fields.Item("UniqueID").Value;
                oForm.Items.Item("Sequence").Specific.Value = oRecordSet.Fields.Item("Sequence").Value;
                oForm.Items.Item("UseYN").Specific.Select(oRecordSet.Fields.Item("UseYN").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_SY015_SAVE
        /// </summary>
        private void PS_SY015_SAVE(int oRow)
        {
            int i;
            string sQry;
            string pGubun;
            string pFSGubun;
            string pUserID;
            string Modual;
            string Sub1;
            string Sub2;
            string Sub3;
            string UserID;
            string No;
            string Level;
            string FatherID;
            string Position;
            string Strings_Renamed;
            string UniqueID;
            string pUniqueID;
            string Sequence;
            string AuthType;
            string beValue;
            string afValue;
            string beCode;
            string UseYN;
            string MType;
            string errMessage = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                pGubun = oForm.Items.Item("pGubun").Specific.Value.ToString().Trim();
                pFSGubun = oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                pUserID = oForm.Items.Item("pUserID").Specific.Value.ToString().Trim();
                Modual = oForm.Items.Item("Modual").Specific.Value.ToString().Trim();
                Sub1 = oForm.Items.Item("Sub1").Specific.Value.ToString().Trim();
                Sub2 = oForm.Items.Item("Sub2").Specific.Value.ToString().Trim();
                Sub3 = oForm.Items.Item("Sub3").Specific.Value.ToString().Trim();
                UserID = oForm.Items.Item("UserID").Specific.Value.ToString().Trim();
                No = oForm.Items.Item("No").Specific.Value.ToString().Trim();
                Level = oForm.Items.Item("Level").Specific.Value.ToString().Trim();
                FatherID = oForm.Items.Item("FatherID").Specific.Value.ToString().Trim();
                Position = oForm.Items.Item("Position").Specific.Value.ToString().Trim();
                Strings_Renamed = oForm.Items.Item("Strings").Specific.Value.ToString().Trim();
                UniqueID = oForm.Items.Item("UniqueID").Specific.Value.ToString().Trim();
                Sequence = oForm.Items.Item("Sequence").Specific.Value.ToString().Trim();
                pUniqueID = oForm.Items.Item("UniqueID").Specific.Value.ToString().Trim();
                UseYN = oForm.Items.Item("UseYN").Specific.Value.ToString().Trim();
                MType = oForm.Items.Item("MType").Specific.Value.ToString().Trim();
                AuthType = "";

                if (oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim()  != "A")
                {
                    sQry = "EXEC PS_SY015_03 '" + pFSGubun + "', '" + pUserID + "', '" + pUniqueID + "', '" + Sequence + "', '" + UniqueID + "', '" + UserID + "', '" + FatherID + "', '" + Strings_Renamed + "', '" + Position + "', '";
                    sQry += Level + "', '" + No + "', '" + pGubun + "', '" + AuthType + "', '" + PSH_Globals.oCompany.UserName + "','" + MType + "','" + UseYN + "'";

                    oDS_PS_SY015A.ExecuteQuery(sQry);
                }
                else
                {
                    if(string.IsNullOrEmpty(oForm.Items.Item("refData").Specific.value))
                    {
                        errMessage = "관련근거 입력은 필수 입니다.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oForm.Items.Item("MType").Specific.value))
                    {
                        errMessage = "변경타입 입력은 필수 입니다.";
                        throw new Exception();
                    }
                    for (i = 0; i < oGrid1.Rows.Count; i++)
                    {
                        if(temp[i] != oGrid1.DataTable.Columns.Item("UniqueID").Cells.Item(i).Value + "_" + oGrid1.DataTable.Columns.Item("UserID").Cells.Item(i).Value + "_" + oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value)
                        {
                            sQry = "EXEC PS_SY015_03 'A', '', '', '', '', '" + oGrid1.DataTable.Columns.Item("UserID").Cells.Item(i).Value + "', '', '', '', '', '";
                            sQry += oGrid1.DataTable.Columns.Item("seq").Cells.Item(i).Value + "', '', '" + oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value + "', '" + PSH_Globals.oCompany.UserName + "','" + MType + "',''";
                            oRecordSet01.DoQuery(sQry);

                            if (oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value == "R") {
                                afValue = "읽기 전용";
                            }
                            else if (oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value == "Y")
                            {
                                afValue = "모든 권한";
                            }
                            else if (oGrid1.DataTable.Columns.Item("AuthType").Cells.Item(i).Value == "N")
                            {
                                afValue = "권한 없음";
                            }
                            else
                            {
                                afValue = "해당 없음";
                            }

                            beCode = codeHelpClass.Right(temp[i], 1);
                            if (beCode == "R")
                            {
                                beValue = "읽기 전용";
                            }
                            else if (beCode == "Y")
                            {
                                beValue = "모든 권한";
                            }
                            else if (beCode == "N")
                            {
                                beValue = "권한 없음";
                            }
                            else
                            {
                                beValue = "해당 없음";
                            }

                            sQry = "insert into PS_SY021 SELECT 'A','" + oGrid1.DataTable.Columns.Item("UserID").Cells.Item(i).Value + "','" + oGrid1.DataTable.Columns.Item("String").Cells.Item(i).Value + "^" + beValue + "','";
                            sQry += afValue + "','" + PSH_Globals.oCompany.UserName + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', '" + oForm.Items.Item("refData").Specific.value + "','" + MType + "'";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                }
                PS_SY015_MTX01();
                PSH_Globals.SBO_Application.MessageBox("입력완료");
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_SAVE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_SY015_TitleSetting(int iRow)
        {
            int i;
            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                for (i = 0; i < oGrid1.DataTable.Columns.Count; i++)
                {
                    switch (oGrid1.Columns.Item(i).TitleObject.Caption)
                    {
                        case "AuthType":
                            oGrid1.Columns.Item("AuthType").Editable = true;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("AuthType");

                            oComboCol.ValidValues.Add("Y", "모든 권한");
                            oComboCol.ValidValues.Add("N", "권한 없음");
                            oComboCol.ValidValues.Add("-", "해당 없음");
                            oComboCol.ValidValues.Add("R", "읽기 전용");

                            oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;

                        default:
                            oGrid1.Columns.Item(oGrid1.Columns.Item(i).TitleObject.Caption).Editable = false;
                            break;
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
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
        /// PS_SY015_Delete
        /// </summary>
        private void PS_SY015_Delete()
        {
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (PSH_Globals.SBO_Application.MessageBox("삭제하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    errMessage = "취소되었습니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("pFSGubun").Specific.Value == "F")
                {
                    sQry = "  delete";
                    sQry += " from    Authority_Folder";
                    sQry += " where   UniqueID = '" + oForm.Items.Item("UniqueID").Specific.Value + "'";

                    oRecordSet01.DoQuery(sQry);
                }
                else
                {
                    sQry = "  delete";
                    sQry += " from    Authority_Screen";
                    sQry += " where   UniqueID = '" + oForm.Items.Item("UniqueID").Specific.Value + "'";

                    oRecordSet01.DoQuery(sQry);

                    sQry = "  delete";
                    sQry += " from    Authority_User";
                    sQry += " where   seq = '" + oForm.Items.Item("Sequence").Specific.Value + "'";

                }
                oRecordSet01.DoQuery(sQry);
                PS_SY015_MTX01();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_Delete_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_SY015_Copy
        /// </summary>
        private void PS_SY015_Copy()
        {
            string sQry;
            string pUserID;
            string CPUserID;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                pUserID = oForm.Items.Item("pUserID").Specific.Value.ToString().Trim();
                CPUserID = oForm.Items.Item("CPUserID").Specific.Value.ToString().Trim();
                
                sQry = "select count(1) from Authority_User where UserID ='" + pUserID + "'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    if (string.IsNullOrEmpty(pUserID.ToString().Trim()) || string.IsNullOrEmpty(CPUserID.ToString().Trim()))
                    {
                        errMessage = "대상 ID와 복제ID는 필수입니다.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oForm.Items.Item("refData").Specific.value))
                    {
                        errMessage = "관련근거 입력은 필수 입니다.";
                        throw new Exception();
                    }

                    sQry = "Insert into PS_SY021 select 'A'";
                    sQry += ", '" + oForm.Items.Item("CPUserID").Specific.Value + "'";
                    sQry += ", b.String  + '^모든 권한'";
                    sQry += ", '권한 회수'";
                    sQry += ", '" + PSH_Globals.oCompany.UserName + "'";
                    sQry += ", Convert(varchar(20), GETDATE(),120)";
                    sQry += ", '" + oForm.Items.Item("refData").Specific.Value + "(복제회수)','"+ oForm.Items.Item("MType").Specific.Value + "'";
                    sQry += " from Authority_User a inner join authority_screen b on a.seq = b.Seq";
                    sQry += " where  a.UserID = '" + oForm.Items.Item("CPUserID").Specific.Value + "'";
                    sQry += "   and a.AuthType in ('Y','R')";
                    oRecordSet.DoQuery(sQry);

                    sQry = "delete from Authority_User where UserID ='" + oForm.Items.Item("CPUserID").Specific.Value + "'";
                    oRecordSet.DoQuery(sQry);

                    sQry = "Insert into Authority_User";
                    sQry += " select seq,'" + CPUserID + "'";
                    sQry += ", 'Y'";
                    sQry += ", AuthType";
                    sQry += ", GETDATE()";
                    sQry += ", '" + PSH_Globals.oCompany.UserName + "'";
                    sQry += "  from Authority_User";
                    sQry += "  where UserID ='" + pUserID + "'";
                    sQry += "    and seq    <> ''";
                    oRecordSet.DoQuery(sQry);

                    sQry = "Insert into PS_SY021 select 'A'";
                    sQry += ", '" + oForm.Items.Item("CPUserID").Specific.Value + "'";
                    sQry += ", b.String  + '^해당 없음'";
                    sQry += ", '모든 권한'";
                    sQry += ", '" + PSH_Globals.oCompany.UserName + "'";
                    sQry += ", Convert(varchar(20), GETDATE(),120)";
                    sQry += ", '" + oForm.Items.Item("refData").Specific.Value + "(복제대상 : " + pUserID + ")','" + oForm.Items.Item("MType").Specific.Value + "'";
                    sQry += " from Authority_User a inner join authority_screen b on a.seq = b.Seq";
                    sQry += " where  a.UserID = '" + oForm.Items.Item("CPUserID").Specific.Value + "'";
                    sQry += "   and a.AuthType in ('Y','R')";
                    oRecordSet.DoQuery(sQry);

                    PSH_Globals.SBO_Application.MessageBox("복제완료!");
                }
                else
                {
                    errMessage = "대상의 권한이 없습니다.";
                    throw new Exception();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_SY015_Copy_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
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
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("처리중....", 0, false);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Find")
                    {
                        PS_SY015_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        PS_SY015_SAVE(pVal.Row);
                    }
                    else if (pVal.ItemUID == "Btn_del")
                    {
                        PS_SY015_Delete();
                    }
                    else if (pVal.ItemUID == "Bt_Copy")
                    {
                        PS_SY015_Copy();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_SY015_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_SY015_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PS_SY015_FormItemEnabled();
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
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
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
                if (pVal.CharPressed == 9)
                {
                    if (pVal.ItemUID == "pUserID")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("pUserID").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "CPUserID")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CPUserID").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "UserID")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("UserID").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "pUnique")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("pUnique").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "pUserID")
                        {
                            oForm.Items.Item("pUserNm").Specific.Value = dataHelpClass.Get_ReData("U_Name", "User_Code", "OUSR", "'" + oForm.Items.Item("pUserID").Specific.Value.ToString().Trim() + "'", "");
                        }
                        else if (pVal.ItemUID == "CPUserID")
                        {
                            oForm.Items.Item("CPUserNm").Specific.Value = dataHelpClass.Get_ReData("U_Name", "User_Code", "OUSR", "'" + oForm.Items.Item("CPUserID").Specific.Value.ToString().Trim() + "'", "");
                        }
                        else if (pVal.ItemUID == "pUnique")
                        {
                            oForm.Items.Item("pUniqNm").Specific.Value = dataHelpClass.Get_ReData("String", "UniqueID", "Authority_Screen", "'" + oForm.Items.Item("pUnique").Specific.Value.ToString().Trim() + "'", "");
                        }
                        else if (pVal.ItemUID == "UserID")
                        {
                            oForm.Items.Item("UserNm").Specific.Value = dataHelpClass.Get_ReData("U_Name", "User_Code", "OUSR", "'" + oForm.Items.Item("UserID").Specific.Value.ToString().Trim() + "'", "");
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemChanged == true)
                    {
                        if(pVal.ItemUID == "pFSGubun")
                        {
                            if (oForm.Items.Item("pFSGubun").Specific.Value == "C")
                            {
                                oForm.Items.Item("MType").Specific.Select("C");
                            }
                            else
                            {
                                oForm.Items.Item("MType").Specific.Select("M");
                            }
                        }
                        else if(pVal.ItemUID == "Sub1" || pVal.ItemUID == "Sub2" || pVal.ItemUID == "Sub3" || pVal.ItemUID == "No")
                        {
                            if (oForm.Items.Item("pFSGubun").Specific.Value == "F")
                            {
                                oForm.DataSources.UserDataSources.Item("UniqueId").Value = oForm.Items.Item("Modual").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub1").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub2").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub3").Specific.Value.ToString().Trim() + oForm.Items.Item("No").Specific.Value.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                                oForm.Items.Item("Sequence").Specific.Value = oForm.Items.Item("Modual").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub1").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub2").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub3").Specific.Value.ToString().Trim() + oForm.Items.Item("No").Specific.Value.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                            }
                            oForm.Items.Item("Sequence").Specific.Value = oForm.Items.Item("Modual").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub1").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub2").Specific.Value.ToString().Trim() + oForm.Items.Item("Sub3").Specific.Value.ToString().Trim() + oForm.Items.Item("No").Specific.Value.ToString().Trim() + oForm.Items.Item("pFSGubun").Specific.Value.ToString().Trim();
                        }
                    }
                }
                PS_SY015_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SY015A);
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
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0) {
                                if (oForm.Items.Item("pFSGubun").Specific.Value != "A")
                                {
                                    PS_SY015_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                }
                            }
                            break;
                    }
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
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
                oForm.Freeze(false);
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
                            PS_SY015_FormItemEnabled();
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PS_SY015_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
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
