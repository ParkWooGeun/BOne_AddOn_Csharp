using System;
using System.Collections.Generic;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 교육계획VS실적조회
    /// </summary>
    internal class PH_PY205 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        public SAPbouiCOM.Grid oGrid01;
        public SAPbouiCOM.Grid oGrid02;
        public SAPbouiCOM.Grid oGrid03;

        public SAPbouiCOM.DataTable oDS_PH_PY205A;
        public SAPbouiCOM.DataTable oDS_PH_PY205B;
        public SAPbouiCOM.DataTable oDS_PH_PY205C;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY205.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY205_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY205");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY205_CreateItems();
                PH_PY205_ComboBox_Setting();
                PH_PY205_EnableMenus();
                oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
                PH_PY205_SetDocument(oFormDocEntry01);
                //Call PH_PY205_FormResize

                oForm.DataSources.UserDataSources.Item("StdYear01").Value = DateTime.Now.ToString("yyyy"); //기준년도
                oForm.DataSources.UserDataSources.Item("CancelYN01").Value = "Y";
                oForm.DataSources.UserDataSources.Item("FrMt02").Value = DateTime.Now.ToString("yyyyMM"); //기준년월(시작)
                oForm.DataSources.UserDataSources.Item("ToMt02").Value = DateTime.Now.ToString("yyyyMM"); //기준년월(종료)
                oForm.DataSources.UserDataSources.Item("Check201").Value = "Y"; //수정계획 미출력
                oForm.DataSources.UserDataSources.Item("StdYear03").Value = DateTime.Now.ToString("yyyy"); //기준년도
                oForm.DataSources.UserDataSources.Item("CancelYN03").Value = "Y";

                oForm.Items.Item("MSTCOD01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        private void PH_PY205_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oGrid02 = oForm.Items.Item("Grid02").Specific;
                oGrid03 = oForm.Items.Item("Grid03").Specific;

                oForm.DataSources.DataTables.Add("PH_PY205A");
                oForm.DataSources.DataTables.Add("PH_PY205B");
                oForm.DataSources.DataTables.Add("PH_PY205C");

                oGrid01.DataTable = oForm.DataSources.DataTables.Item("PH_PY205A");
                oGrid02.DataTable = oForm.DataSources.DataTables.Item("PH_PY205B");
                oGrid03.DataTable = oForm.DataSources.DataTables.Item("PH_PY205C");

                oDS_PH_PY205A = oForm.DataSources.DataTables.Item("PH_PY205A");
                oDS_PH_PY205B = oForm.DataSources.DataTables.Item("PH_PY205B");
                oDS_PH_PY205C = oForm.DataSources.DataTables.Item("PH_PY205C");

                ////////////Folder01//////////_S
                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD01").Specific.DataBind.SetBound(true, "", "CLTCOD01");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode01").Specific.DataBind.SetBound(true, "", "TeamCode01");

                //재직여부
                oForm.DataSources.UserDataSources.Add("Status01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Status01").Specific.DataBind.SetBound(true, "", "Status01");

                //사원번호
                oForm.DataSources.UserDataSources.Add("MSTCOD01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTCOD01").Specific.DataBind.SetBound(true, "", "MSTCOD01");

                //사원성명
                oForm.DataSources.UserDataSources.Add("MSTNAM01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("MSTNAM01").Specific.DataBind.SetBound(true, "", "MSTNAM01");

                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear01").Specific.DataBind.SetBound(true, "", "StdYear01");

                //기준월
                oForm.DataSources.UserDataSources.Add("StdMt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdMt01").Specific.DataBind.SetBound(true, "", "StdMt01");

                //계획구분
                oForm.DataSources.UserDataSources.Add("PlnCls01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PlnCls01").Specific.DataBind.SetBound(true, "", "PlnCls01");

                //상태
                oForm.DataSources.UserDataSources.Add("DocSts01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DocSts01").Specific.DataBind.SetBound(true, "", "DocSts01");

                //교육명
                oForm.DataSources.UserDataSources.Add("EduName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("EduName01").Specific.DataBind.SetBound(true, "", "EduName01");

                //교육기관
                oForm.DataSources.UserDataSources.Add("EduOrg01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("EduOrg01").Specific.DataBind.SetBound(true, "", "EduOrg01");

                //취소제외
                oForm.DataSources.UserDataSources.Add("CancelYN01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("CancelYN01").Specific.DataBind.SetBound(true, "", "CancelYN01");
                ////////////Folder01//////////_E

                ////////////Folder02//////////_S
                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD02").Specific.DataBind.SetBound(true, "", "CLTCOD02");

                //기준년월(시작)
                oForm.DataSources.UserDataSources.Add("FrMt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("FrMt02").Specific.DataBind.SetBound(true, "", "FrMt02");

                //기준년월(종료)
                oForm.DataSources.UserDataSources.Add("ToMt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("ToMt02").Specific.DataBind.SetBound(true, "", "ToMt02");

                //재직여부
                oForm.DataSources.UserDataSources.Add("Status02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Status02").Specific.DataBind.SetBound(true, "", "Status02");

                //수정계획 미출력
                oForm.DataSources.UserDataSources.Add("Check201", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Check201").Specific.DataBind.SetBound(true, "", "Check201");
                ////////////Folder02//////////_E

                ////////////Folder03//////////_S
                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD03").Specific.DataBind.SetBound(true, "", "CLTCOD03");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode03").Specific.DataBind.SetBound(true, "", "TeamCode03");

                //재직여부
                oForm.DataSources.UserDataSources.Add("Status03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Status03").Specific.DataBind.SetBound(true, "", "Status03");

                //사원번호
                oForm.DataSources.UserDataSources.Add("MSTCOD03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTCOD03").Specific.DataBind.SetBound(true, "", "MSTCOD03");

                //사원성명
                oForm.DataSources.UserDataSources.Add("MSTNAM03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("MSTNAM03").Specific.DataBind.SetBound(true, "", "MSTNAM03");

                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear03").Specific.DataBind.SetBound(true, "", "StdYear03");

                //기준월
                oForm.DataSources.UserDataSources.Add("StdMt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdMt03").Specific.DataBind.SetBound(true, "", "StdMt03");

                //계획구분
                oForm.DataSources.UserDataSources.Add("PlnCls03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PlnCls03").Specific.DataBind.SetBound(true, "", "PlnCls03");

                //상태
                oForm.DataSources.UserDataSources.Add("DocSts03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DocSts03").Specific.DataBind.SetBound(true, "", "DocSts03");

                //교육명
                oForm.DataSources.UserDataSources.Add("EduName03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("EduName03").Specific.DataBind.SetBound(true, "", "EduName03");

                //교육기관
                oForm.DataSources.UserDataSources.Add("EduOrg03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("EduOrg03").Specific.DataBind.SetBound(true, "", "EduOrg03");

                //취소제외
                oForm.DataSources.UserDataSources.Add("CancelYN03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("CancelYN03").Specific.DataBind.SetBound(true, "", "CancelYN03");
                ////////////Folder03//////////_E
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY205_ComboBox_Setting()
        {
            //SAPbouiCOM.ComboBox oCombo = null;
            string sQry = string.Empty;
            
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                ////////////Folder01//////////_S
                //재직여부
                oForm.Items.Item("Status01").Specific.ValidValues.Add("1", "재직자");
                oForm.Items.Item("Status01").Specific.ValidValues.Add("2", "퇴직자포함");
                oForm.Items.Item("Status01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //계획구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm AS [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P239'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";

                oForm.Items.Item("PlnCls01").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("PlnCls01").Specific, sQry, "", false, false);
                oForm.Items.Item("PlnCls01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //기준월
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm AS [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = '4'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";

                oForm.Items.Item("StdMt01").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("StdMt01").Specific, sQry, "", false, false);
                oForm.Items.Item("StdMt01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //상태
                oForm.Items.Item("DocSts01").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("DocSts01").Specific.ValidValues.Add("O", "계획");
                oForm.Items.Item("DocSts01").Specific.ValidValues.Add("P", "완료");
                oForm.Items.Item("DocSts01").Specific.ValidValues.Add("C", "취소");
                oForm.Items.Item("DocSts01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                ////////////Folder01//////////_E

                ////////////Folder02//////////_S
                //재직여부
                oForm.Items.Item("Status02").Specific.ValidValues.Add("1", "재직자");
                oForm.Items.Item("Status02").Specific.ValidValues.Add("2", "퇴직자포함");
                oForm.Items.Item("Status02").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                ////////////Folder02//////////_E

                ////////////Folder03//////////_S
                //재직여부
                oForm.Items.Item("Status03").Specific.ValidValues.Add("1", "재직자");
                oForm.Items.Item("Status03").Specific.ValidValues.Add("2", "퇴직자포함");
                oForm.Items.Item("Status03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //계획구분
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm AS [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = 'P239'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";

                oForm.Items.Item("PlnCls03").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("PlnCls03").Specific, sQry, "", false, false);
                oForm.Items.Item("PlnCls03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //월
                sQry = "        SELECT      U_Code AS [Code],";
                sQry = sQry + "             U_CodeNm AS [Name]";
                sQry = sQry + " FROM        [@PS_HR200L]";
                sQry = sQry + " WHERE       Code = '4'";
                sQry = sQry + "             AND U_UseYN = 'Y'";
                sQry = sQry + " ORDER BY    U_Seq";

                oForm.Items.Item("StdMt03").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("StdMt03").Specific, sQry, "", false, false);
                oForm.Items.Item("StdMt03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //상태
                oForm.Items.Item("DocSts03").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("DocSts03").Specific.ValidValues.Add("O", "계획");
                oForm.Items.Item("DocSts03").Specific.ValidValues.Add("P", "완료");
                oForm.Items.Item("DocSts03").Specific.ValidValues.Add("C", "취소");
                oForm.Items.Item("DocSts03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                ////////////Folder03//////////_E

                ////////////매트릭스//////////_S
                //    '상태
                //    Call oMat01.Columns("DocStatus").ValidValues.Add("O", "계획")
                //    Call oMat01.Columns("DocStatus").ValidValues.Add("P", "완료")
                //    Call oMat01.Columns("DocStatus").ValidValues.Add("C", "취소")
                //
                //    '계획구분
                //    sQry = "            SELECT      U_Code AS [Code],"
                //    sQry = sQry & "                 U_CodeNm AS [Name]"
                //    sQry = sQry & " FROM        [@PS_HR200L]"
                //    sQry = sQry & " WHERE       Code = 'P239'"
                //    sQry = sQry & "                 AND U_UseYN = 'Y'"
                //    sQry = sQry & " ORDER BY  U_Seq"
                //    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PlnCls"), sQry)
                //
                //    '이수(수료)증
                //    Call oMat01.Columns("Certi").ValidValues.Add("O", "미제출")
                //    Call oMat01.Columns("Certi").ValidValues.Add("C", "제출")
                //
                //    '보고서
                //    Call oMat01.Columns("Report").ValidValues.Add("O", "미제출")
                //    Call oMat01.Columns("Report").ValidValues.Add("C", "제출")
                ////////////매트릭스//////////_E
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY205_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY205_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY205_FormItemEnabled();
                }
                else
                {
                    //oForm.Mode = fm_FIND_MODE
                    PH_PY205_FormItemEnabled();
                    //oForm.Items("Code").Specific.Value = oFormDocEntry01
                    //oForm.Items("1").CLICK ct_Regular
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY205_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD01", true);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD02", true);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD03", true);

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD01", true);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD02", true);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD03", true);

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD01", false);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD02", false);
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD03", false);

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PH_PY205_FormResize()
        {
            try
            {
                //그룹박스 크기 동적 할당
                oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 145;
                oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

                if (oGrid01.Columns.Count > 0)
                {
                    oGrid01.AutoResizeColumns();
                }

                if (oGrid02.Columns.Count > 0)
                {
                    oGrid02.AutoResizeColumns();
                }

                if (oGrid03.Columns.Count > 0)
                {
                    oGrid03.AutoResizeColumns();
                }

                //    If oGrid04.Columns.Count > 0 Then
                //        oGrid04.AutoResizeColumns
                //    End If
                //
                //    If oGrid05.Columns.Count > 0 Then
                //        oGrid05.AutoResizeColumns
                //    End If
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_FormResize_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY205_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int loopCount = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
                        
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                switch (oUID)
                {
                    case "CLTCOD01":

                        CLTCOD = oForm.Items.Item("CLTCOD01").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("TeamCode01").Specific.ValidValues.Count > 0)
                        {
                            for (loopCount = oForm.Items.Item("TeamCode01").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                            {
                                oForm.Items.Item("TeamCode01").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //부서콤보세팅
                        oForm.Items.Item("TeamCode01").Specific.ValidValues.Add("%", "전체");
                        sQry = "        SELECT      U_Code AS [Code],";
                        sQry = sQry + "             U_CodeNm As [Name]";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '1'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + "             AND U_Char2 = '" + CLTCOD + "'";
                        sQry = sQry + " UNION ALL ";
                        sQry = sQry + " SELECT      '9999' AS [Code],";
                        sQry = sQry + "             '전부서' As [Name]";
                        sQry = sQry + " ORDER BY    U_Code";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode01").Specific, sQry, "", false, false);
                        oForm.Items.Item("TeamCode01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "MSTCOD01":

                        oForm.Items.Item("MSTNAM01").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD01").Specific.Value + "'", ""); //성명
                        break;

                    case "CLTCOD03":

                        CLTCOD = oForm.Items.Item("CLTCOD03").Specific.Value.ToString().Trim();

                        if (oForm.Items.Item("TeamCode03").Specific.ValidValues.Count > 0)
                        {
                            for (loopCount = oForm.Items.Item("TeamCode03").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
                            {
                                oForm.Items.Item("TeamCode03").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        //부서콤보세팅
                        oForm.Items.Item("TeamCode03").Specific.ValidValues.Add("%", "전체");
                        sQry = "        SELECT      U_Code AS [Code],";
                        sQry = sQry + "             U_CodeNm As [Name]";
                        sQry = sQry + " FROM        [@PS_HR200L]";
                        sQry = sQry + " WHERE       Code = '1'";
                        sQry = sQry + "             AND U_UseYN = 'Y'";
                        sQry = sQry + "             AND U_Char2 = '" + CLTCOD + "'";
                        sQry = sQry + " ORDER BY    U_Seq";
                        dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode03").Specific, sQry, "", false, false);
                        oForm.Items.Item("TeamCode03").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        break;

                    case "MSTCOD03":

                        oForm.Items.Item("MSTNAM03").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD03").Specific.Value + "'", ""); //성명
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// FormClear
        /// </summary>
        private void PH_PY205_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY205'", "");

                if (Convert.ToDouble(DocEntry) == 0)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FolderClear
        /// </summary>
        private void PH_PY205_FolderClear()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.PaneLevel == 1)
                {
                    if (oGrid01.Columns.Count == 0)
                    {
                        oForm.Items.Item("MSTCOD01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else if (oForm.PaneLevel == 2)
                {
                    if (oGrid02.Columns.Count == 0)
                    {
                        oForm.Items.Item("FrMt02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else if (oForm.PaneLevel == 3)
                {
                    if (oGrid03.Columns.Count == 0)
                    {
                        PH_PY205_FlushToItemValue("CLTCOD03", 0, "");
                        oForm.Items.Item("MSTCOD03").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //사업장 콤보박스 이벤트 강제 실행(하위 팀 콤보박스의 Binding 재설정을 위함)
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_FolderClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private bool PH_PY205_DataValidCheck()
        {
            bool functionReturnValue = false;
            functionReturnValue = false;

            short ErrNum = 0;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD01").Specific.Value))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                if (ErrNum == 1)
                {
                    functionReturnValue = false;
                    oForm.Items.Item("CLTCOD01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_DataValidCheck_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY205_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY205A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
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
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    functionReturnValue = false;
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);    
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 교육계획상태(개인별상세) 조회
        /// </summary>
        private void PH_PY205_DataFind01()
        {
            string sQry = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string MSTCOD = string.Empty; //사번
            string StdYear = string.Empty; //기준년도
            string StdMt = string.Empty; //기준월
            string Status = string.Empty; //재직여부
            string DocStatus = string.Empty; //상태
            string PlnCls = string.Empty; //계획구분
            string EduName = string.Empty; //교육명
            string EduOrg = string.Empty; //교육기관
            string CancelYN = string.Empty; //취소제외

            CLTCOD = oForm.Items.Item("CLTCOD01").Specific.Value.ToString().Trim(); //사업장
            TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim(); //부서
            MSTCOD = oForm.Items.Item("MSTCOD01").Specific.Value.ToString().Trim(); //사번
            StdYear = oForm.Items.Item("StdYear01").Specific.Value.ToString().Trim(); //기준년도
            StdMt = oForm.Items.Item("StdMt01").Specific.Value.ToString().Trim(); //기준월
            Status = oForm.Items.Item("Status01").Specific.Value.ToString().Trim(); //재직여부
            DocStatus = oForm.Items.Item("DocSts01").Specific.Value.ToString().Trim(); //상태
            PlnCls = oForm.Items.Item("PlnCls01").Specific.Value.ToString().Trim(); //계획구분
            EduName = oForm.Items.Item("EduName01").Specific.Value.ToString().Trim(); //교육명
            EduOrg = oForm.Items.Item("EduOrg01").Specific.Value.ToString().Trim(); //교육기관

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

            try
            {
                oForm.Freeze(true);

                if (oForm.DataSources.UserDataSources.Item("CancelYN01").Value == "Y") //취소제외 여부
                {
                    CancelYN = "Y";
                }
                else
                {
                    CancelYN = "N";
                }

                sQry = "      EXEC [PH_PY205_01] '";
                sQry = sQry + CLTCOD + "','"; //사업장
                sQry = sQry + TeamCode + "','"; //부서
                sQry = sQry + MSTCOD + "','"; //사번
                sQry = sQry + StdYear + "','"; //기준년도
                sQry = sQry + StdMt + "','"; //기준월
                sQry = sQry + Status + "','"; //재직여부
                sQry = sQry + DocStatus + "','"; //상태
                sQry = sQry + PlnCls + "','"; //계획구분
                sQry = sQry + EduName + "','"; //교육명
                sQry = sQry + EduOrg + "','"; //교육기관
                sQry = sQry + CancelYN + "'"; //취소제외

                oDS_PH_PY205A.ExecuteQuery(sQry);

                oGrid01.Columns.Item(12).RightJustified = true;
                oGrid01.Columns.Item(13).RightJustified = true;
                //    oGrid01.Columns(14).RightJustified = True
                //    oGrid01.Columns(15).RightJustified = True
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_DataFind01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// 계획VS실적(팀별집계) 조회
        /// </summary>
        private void PH_PY205_DataFind02()
        {
            string sQry = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string FrMt = string.Empty; //기준년월(시작)
            string ToMt = string.Empty; //기준년월(종료)
            string Status = string.Empty; //재직여부

            CLTCOD = oForm.Items.Item("CLTCOD02").Specific.Value.ToString().Trim(); //사업장
            FrMt = oForm.Items.Item("FrMt02").Specific.Value.ToString().Trim(); //기준년월(시작)
            ToMt = oForm.Items.Item("ToMt02").Specific.Value.ToString().Trim(); //기준년월(종료)
            Status = oForm.Items.Item("Status02").Specific.Value.ToString().Trim(); //재직여부

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

            try
            {
                oForm.Freeze(true);

                sQry = "      EXEC [PH_PY205_03] '";
                sQry = sQry + CLTCOD + "','"; //사업장
                sQry = sQry + FrMt + "','"; //기준년월(시작)
                sQry = sQry + ToMt + "','"; //기준년월(종료)
                sQry = sQry + Status + "'"; //재직여부

                oDS_PH_PY205B.ExecuteQuery(sQry);

                oGrid02.Columns.Item(2).RightJustified = true;
                oGrid02.Columns.Item(3).RightJustified = true;
                oGrid02.Columns.Item(4).RightJustified = true;
                oGrid02.Columns.Item(5).RightJustified = true;
                oGrid02.Columns.Item(6).RightJustified = true;
                oGrid02.Columns.Item(7).RightJustified = true;
                oGrid02.Columns.Item(8).RightJustified = true;
                oGrid02.Columns.Item(9).RightJustified = true;
                oGrid02.Columns.Item(10).RightJustified = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_DataFind02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// 계획VS실적(교육별, 개인별) 조회
        /// </summary>
        private void PH_PY205_DataFind03()
        {
            string sQry = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string MSTCOD = string.Empty; //사번
            string StdYear = string.Empty; //기준년도
            string StdMt = string.Empty; //기준월
            string Status = string.Empty; //재직여부
            string DocStatus = string.Empty; //상태
            string PlnCls = string.Empty; //계획구분
            string EduName = string.Empty; //교육명
            string EduOrg = string.Empty; //교육기관
            string CancelYN = string.Empty; //취소제외

            CLTCOD = oForm.Items.Item("CLTCOD03").Specific.Value.ToString().Trim(); //사업장
            TeamCode = oForm.Items.Item("TeamCode03").Specific.Value.ToString().Trim(); //부서
            MSTCOD = oForm.Items.Item("MSTCOD03").Specific.Value.ToString().Trim(); //사번
            StdYear = oForm.Items.Item("StdYear03").Specific.Value.ToString().Trim(); //기준년도
            StdMt = oForm.Items.Item("StdMt03").Specific.Value.ToString().Trim(); //기준월
            Status = oForm.Items.Item("Status03").Specific.Value.ToString().Trim(); //재직여부
            DocStatus = oForm.Items.Item("DocSts03").Specific.Value.ToString().Trim(); //상태
            PlnCls = oForm.Items.Item("PlnCls03").Specific.Value.ToString().Trim(); //계획구분
            EduName = oForm.Items.Item("EduName03").Specific.Value.ToString().Trim(); //교육명
            EduOrg = oForm.Items.Item("EduOrg03").Specific.Value.ToString().Trim(); //교육기관

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

            try
            {
                oForm.Freeze(true);

                if (oForm.DataSources.UserDataSources.Item("CancelYN03").Value == "Y") //취소제외 여부
                {
                    CancelYN = "Y";
                }
                else
                {
                    CancelYN = "N";
                }

                sQry = "      EXEC [PH_PY205_05] '";
                sQry = sQry + CLTCOD + "','"; //사업장
                sQry = sQry + TeamCode + "','"; //부서
                sQry = sQry + MSTCOD + "','"; //사번
                sQry = sQry + StdYear + "','"; //기준년도
                sQry = sQry + StdMt + "','"; //기준월
                sQry = sQry + Status + "','"; //재직여부
                sQry = sQry + DocStatus + "','"; //상태
                sQry = sQry + PlnCls + "','"; //계획구분
                sQry = sQry + EduName + "','"; //교육명
                sQry = sQry + EduOrg + "','"; //교육기관
                sQry = sQry + CancelYN + "'"; //취소제외

                oDS_PH_PY205C.ExecuteQuery(sQry);

                oGrid03.Columns.Item(11).RightJustified = true;
                oGrid03.Columns.Item(12).RightJustified = true;
                oGrid03.Columns.Item(13).RightJustified = true;
                oGrid03.Columns.Item(14).RightJustified = true;
                oGrid03.Columns.Item(15).RightJustified = true;
                oGrid03.Columns.Item(16).RightJustified = true;
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_DataFind03_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// 교육계획상태(개인별상세) 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY205_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string MSTCOD = string.Empty; //사번
            string StdYear = string.Empty; //기준년도
            string StdMt = string.Empty; //기준월
            string Status = string.Empty; //재직여부
            string DocStatus = string.Empty; //상태
            string PlnCls = string.Empty; //계획구분
            string EduName = string.Empty; //교육명
            string EduOrg = string.Empty; //교육기관
            string CancelYN = string.Empty; //취소제외

            CLTCOD = oForm.Items.Item("CLTCOD01").Specific.Value.ToString().Trim(); //사업장
            TeamCode = oForm.Items.Item("TeamCode01").Specific.Value.ToString().Trim(); //부서
            MSTCOD = oForm.Items.Item("MSTCOD01").Specific.Value.ToString().Trim(); //사번
            StdYear = oForm.Items.Item("StdYear01").Specific.Value.ToString().Trim(); //기준년도
            StdMt = oForm.Items.Item("StdMt01").Specific.Value.ToString().Trim(); //기준월
            Status = oForm.Items.Item("Status01").Specific.Value.ToString().Trim(); //재직여부
            DocStatus = oForm.Items.Item("DocSts01").Specific.Value.ToString().Trim(); //상태
            PlnCls = oForm.Items.Item("PlnCls01").Specific.Value.ToString().Trim(); //계획구분
            EduName = oForm.Items.Item("EduName01").Specific.Value.ToString().Trim(); //교육명
            EduOrg = oForm.Items.Item("EduOrg01").Specific.Value.ToString().Trim(); //교육기관

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.DataSources.UserDataSources.Item("CancelYN01").Value == "Y") //취소제외 여부
                {
                    CancelYN = "Y";
                }
                else
                {
                    CancelYN = "N";
                }

                WinTitle = "[PH_PY205] 레포트";
                ReportName = "PH_PY205_01.rpt";
                //프로시저 : PH_PY205_02

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD_", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode_", TeamCode)); //팀코드
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD_", MSTCOD)); //사번
                dataPackParameter.Add(new PSH_DataPackClass("@StdYear_", StdYear)); //기준년도
                dataPackParameter.Add(new PSH_DataPackClass("@StdMt_", StdMt)); //기준월
                dataPackParameter.Add(new PSH_DataPackClass("@Status_", Status)); //재직여부
                dataPackParameter.Add(new PSH_DataPackClass("@DocStatus_", DocStatus)); //상태
                dataPackParameter.Add(new PSH_DataPackClass("@PlnCls_", PlnCls)); //계획구분
                dataPackParameter.Add(new PSH_DataPackClass("@EduName_", EduName)); //교육명
                dataPackParameter.Add(new PSH_DataPackClass("@EduOrg_", EduOrg)); //교육기관
                dataPackParameter.Add(new PSH_DataPackClass("@CancelYN_", CancelYN)); //취소제외여부

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 계획VS실적(팀별집계) 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY205_Print_Report02()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;
            
            string CLTCOD = string.Empty; //사업장
            string FrMt = string.Empty; //기준년월(시작)
            string ToMt = string.Empty; //기준년월(종료)
            string Status = string.Empty; //재직여부

            CLTCOD = oForm.Items.Item("CLTCOD02").Specific.Value.ToString().Trim(); //사업장
            FrMt = oForm.Items.Item("FrMt02").Specific.Value.ToString().Trim(); //기준년월(시작)
            ToMt = oForm.Items.Item("ToMt02").Specific.Value.ToString().Trim(); //기준년월(종료)
            Status = oForm.Items.Item("Status02").Specific.Value.ToString().Trim(); //재직여부

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.DataSources.UserDataSources.Item("Check201").Value == "Y") //수정계획 미출력
                {
                    ReportName = "PH_PY205_04.rpt";
                }
                else
                {
                    ReportName = "PH_PY205_02.rpt";
                }
                WinTitle = "[PH_PY205] 레포트";
                //프로시저 : PH_PY205_04

                //sQry = "      EXEC [PH_PY205_04] '";
                //sQry = sQry + CLTCOD + "','"; //사업장
                //sQry = sQry + FrMt + "','"; //기준년월(시작)
                //sQry = sQry + ToMt + "','"; //기준년월(종료)
                //sQry = sQry + Status + "'";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD_", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@FrMt_", FrMt)); //기준년월(시작)
                dataPackParameter.Add(new PSH_DataPackClass("@ToMt_", ToMt)); //기준년월(종료)
                dataPackParameter.Add(new PSH_DataPackClass("@Status_", Status)); //재직여부

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_Print_Report02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 계획VS실적(교육별, 개인별) 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY205_Print_Report03()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string MSTCOD = string.Empty; //사번
            string StdYear = string.Empty; //기준년도
            string StdMt = string.Empty; //기준월
            string Status = string.Empty; //재직여부
            string DocStatus = string.Empty; //상태
            string PlnCls = string.Empty; //계획구분
            string EduName = string.Empty; //교육명
            string EduOrg = string.Empty; //교육기관
            string CancelYN = string.Empty; //취소제외

            CLTCOD = oForm.Items.Item("CLTCOD03").Specific.Value.ToString().Trim(); //사업장
            TeamCode = oForm.Items.Item("TeamCode03").Specific.Value.ToString().Trim(); //부서
            MSTCOD = oForm.Items.Item("MSTCOD03").Specific.Value.ToString().Trim(); //사번
            StdYear = oForm.Items.Item("StdYear03").Specific.Value.ToString().Trim(); //기준년도
            StdMt = oForm.Items.Item("StdMt03").Specific.Value.ToString().Trim(); //기준월
            Status = oForm.Items.Item("Status03").Specific.Value.ToString().Trim(); //재직여부
            DocStatus = oForm.Items.Item("DocSts03").Specific.Value.ToString().Trim(); //상태
            PlnCls = oForm.Items.Item("PlnCls03").Specific.Value.ToString().Trim(); //계획구분
            EduName = oForm.Items.Item("EduName03").Specific.Value.ToString().Trim(); //교육명
            EduOrg = oForm.Items.Item("EduOrg03").Specific.Value.ToString().Trim(); //교육기관

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.DataSources.UserDataSources.Item("CancelYN03").Value == "Y") //취소제외 여부
                {
                    CancelYN = "Y";
                }
                else
                {
                    CancelYN = "N";
                }

                WinTitle = "[PH_PY205] 레포트";
                ReportName = "PH_PY205_03.rpt";
                //프로시저 : PH_PY205_06

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD_", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode_", TeamCode)); //부서
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD_", MSTCOD)); //사번
                dataPackParameter.Add(new PSH_DataPackClass("@StdYear_", StdYear)); //기준년도
                dataPackParameter.Add(new PSH_DataPackClass("@StdMt_", StdMt)); //기준월
                dataPackParameter.Add(new PSH_DataPackClass("@Status_", Status)); //재직여부
                dataPackParameter.Add(new PSH_DataPackClass("@DocStatus_", DocStatus)); //상태
                dataPackParameter.Add(new PSH_DataPackClass("@PlnCls_", PlnCls)); //계획구분
                dataPackParameter.Add(new PSH_DataPackClass("@EduName_", EduName)); //교육명
                dataPackParameter.Add(new PSH_DataPackClass("@EduOrg_", EduOrg)); //교육기관
                dataPackParameter.Add(new PSH_DataPackClass("@CancelYN_", CancelYN)); //취소제외여부

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY205_Print_Report03_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
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

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
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
                    if (pVal.ItemUID == "BtnSrch01")
                    {
                        PH_PY205_DataFind01();
                    }
                    else if (pVal.ItemUID == "BtnSrch02")
                    {
                        PH_PY205_DataFind02();
                    }
                    else if (pVal.ItemUID == "BtnSrch03")
                    {
                        PH_PY205_DataFind03();
                    }
                    else if (pVal.ItemUID == "BtnPrt01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY205_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "BtnPrt02")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY205_Print_Report02);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "BtnPrt03")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY205_Print_Report03);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    //폴더를 사용할 때는 필수 소스_S
                    //Folder01이 선택되었을 때
                    if (pVal.ItemUID == "Folder01")
                    {
                        oForm.PaneLevel = 1;
                        oForm.DefButton = "BtnSrch01";

                        PH_PY205_FolderClear();
                    }

                    //Folder02가 선택되었을 때
                    if (pVal.ItemUID == "Folder02")
                    {
                        oForm.PaneLevel = 2;
                        oForm.DefButton = "BtnSrch02";

                        PH_PY205_FolderClear();
                    }

                    //Folder03가 선택되었을 때
                    if (pVal.ItemUID == "Folder03")
                    {
                        oForm.PaneLevel = 3;
                        oForm.DefButton = "BtnSrch03";

                        PH_PY205_FolderClear();
                        //PH_PY205_FlushToItemValue("CLTCOD03") '사업장 콤보박스 이벤트 강제 실행(하위 팀 콤보박스의 Binding 재설정을 위함)
                    }
                    //폴더를 사용할 때는 필수 소스_E
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD01", ""); //Header(Folder01)-사번
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD03", ""); //Header(Folder02)-사번
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
                        PH_PY205_FlushToItemValue(pVal.ItemUID, 0, "");
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
                        case "Grid01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        case "Grid02":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        case "Grid03":
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        PH_PY205_FlushToItemValue(pVal.ItemUID, 0, "");
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY205_FormItemEnabled();
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
                    SubMain.Remove_Forms(oFormUniqueID01);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY205A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY205B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY205C);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                            PH_PY205_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        //  break;
                        case "1281": //문서찾기
                            PH_PY205_FormItemEnabled();
                            break;
                        case "1282": //문서추가
                            PH_PY205_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PH_PY205_FormItemEnabled();
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

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
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