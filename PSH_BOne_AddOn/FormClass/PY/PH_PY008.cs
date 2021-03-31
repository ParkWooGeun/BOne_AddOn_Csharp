using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{ 
    /// <summary>
    /// 일근태등록
    /// </summary>
    internal class PH_PY008 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Grid oGrid1;
        private SAPbouiCOM.DataTable oDS_PH_PY008;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;
        private string oWorkType;
        private string oRspCodeYN; //담당수정 권한

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY008.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY008_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY008");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "code"

                oForm.Freeze(true);
                PH_PY008_CreateItems();
                PH_PY008_EnableMenus();
                PH_PY008_FormItemEnabled();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PH_PY008_CreateItems()
        {
            string sQry;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;

                oForm.DataSources.DataTables.Add("PH_PY008");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY008");
                oDS_PH_PY008 = oForm.DataSources.DataTables.Item("PH_PY008");
                oGrid1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

                oForm.Items.Item("SCLTCOD").DisplayDesc = true;

                //////////조회//////////_S
                //사업장
                oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

                //일자
                oForm.DataSources.UserDataSources.Add("SPosDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("SPosDate").Specific.DataBind.SetBound(true, "", "SPosDate");

                //요일
                oForm.DataSources.UserDataSources.Add("SDay", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDay").Specific.DataBind.SetBound(true, "", "SDay");

                //평일/휴일
                oForm.DataSources.UserDataSources.Add("SDayOff", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SDayOff").Specific.DataBind.SetBound(true, "", "SDayOff");

                //부서
                oForm.DataSources.UserDataSources.Add("STeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("STeamCode").Specific.DataBind.SetBound(true, "", "STeamCode");

                //담당
                oForm.DataSources.UserDataSources.Add("SRspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("SRspCode").Specific.DataBind.SetBound(true, "", "SRspCode");

                //반
                oForm.DataSources.UserDataSources.Add("SClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("SClsCode").Specific.DataBind.SetBound(true, "", "SClsCode");

                //근무형태
                oForm.DataSources.UserDataSources.Add("SShiftDat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SShiftDat").Specific.DataBind.SetBound(true, "", "SShiftDat");

                //근무형태명
                oForm.DataSources.UserDataSources.Add("ShiftDatNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ShiftDatNm").Specific.DataBind.SetBound(true, "", "ShiftDatNm");

                //근무조
                oForm.DataSources.UserDataSources.Add("SGNMUJO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SGNMUJO").Specific.DataBind.SetBound(true, "", "SGNMUJO");

                //근무조명
                oForm.DataSources.UserDataSources.Add("SGNMUJONm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SGNMUJONm").Specific.DataBind.SetBound(true, "", "SGNMUJONm");

                //사번
                oForm.DataSources.UserDataSources.Add("SMSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SMSTCOD").Specific.DataBind.SetBound(true, "", "SMSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("SFullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SFullName").Specific.DataBind.SetBound(true, "", "SFullName");

                //1조
                oForm.DataSources.UserDataSources.Add("Team1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Team1").Specific.DataBind.SetBound(true, "", "Team1");

                //2조
                oForm.DataSources.UserDataSources.Add("Team2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Team2").Specific.DataBind.SetBound(true, "", "Team2");

                //3조
                oForm.DataSources.UserDataSources.Add("Team3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Team3").Specific.DataBind.SetBound(true, "", "Team3");

                //체크박스(근태이상자보기)
                oForm.DataSources.UserDataSources.Add("Chk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Chk").Specific.ValOn = "Y";
                oForm.Items.Item("Chk").Specific.ValOff = "N";
                oForm.Items.Item("Chk").Specific.DataBind.SetBound(true, "", "Chk");
                oForm.DataSources.UserDataSources.Item("Chk").Value = "N";
                //////////조회//////////_E

                //////////저장//////////_S
                //일자
                oForm.DataSources.UserDataSources.Add("PosDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("PosDate").Specific.DataBind.SetBound(true, "", "PosDate");

                //교대근무인정
                oForm.DataSources.UserDataSources.Add("RotateYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("RotateYN").Specific.DataBind.SetBound(true, "", "RotateYN");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                //반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                //근태구분
                oForm.DataSources.UserDataSources.Add("WorkType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("WorkType").Specific.DataBind.SetBound(true, "", "WorkType");

                //교대일수
                oForm.DataSources.UserDataSources.Add("Rotation", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Rotation").Specific.DataBind.SetBound(true, "", "Rotation");

                //근무형태
                oForm.DataSources.UserDataSources.Add("ShiftDat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ShiftDat").Specific.DataBind.SetBound(true, "", "ShiftDat");

                //근무조
                oForm.DataSources.UserDataSources.Add("GNMUJO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("GNMUJO").Specific.DataBind.SetBound(true, "", "GNMUJO");

                // 비고(2013.11.19 송명규 추가)
                oForm.DataSources.UserDataSources.Add("Comment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("Comment").Specific.DataBind.SetBound(true, "", "Comment");

                //사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                //기본
                oForm.DataSources.UserDataSources.Add("Base", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("Base").Specific.DataBind.SetBound(true, "", "Base");

                //연장
                oForm.DataSources.UserDataSources.Add("Extend", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("Extend").Specific.DataBind.SetBound(true, "", "Extend");

                //심야
                oForm.DataSources.UserDataSources.Add("Midnight", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("Midnight").Specific.DataBind.SetBound(true, "", "Midnight");

                //조출
                oForm.DataSources.UserDataSources.Add("EarlyTo", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("EarlyTo").Specific.DataBind.SetBound(true, "", "EarlyTo");

                //출근
                oForm.DataSources.UserDataSources.Add("GetDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("GetDate").Specific.DataBind.SetBound(true, "", "GetDate");

                oForm.DataSources.UserDataSources.Add("GetTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("GetTime").Specific.DataBind.SetBound(true, "", "GetTime");

                oForm.DataSources.UserDataSources.Add("Day1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Day1").Specific.DataBind.SetBound(true, "", "Day1");

                oForm.DataSources.UserDataSources.Add("DayOff1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DayOff1").Specific.DataBind.SetBound(true, "", "DayOff1");

                //특근
                oForm.DataSources.UserDataSources.Add("Special", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("Special").Specific.DataBind.SetBound(true, "", "Special");

                //특연
                oForm.DataSources.UserDataSources.Add("SpExtend", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("SpExtend").Specific.DataBind.SetBound(true, "", "SpExtend");

                //교육훈련
                oForm.DataSources.UserDataSources.Add("EducTran", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("EducTran").Specific.DataBind.SetBound(true, "", "EducTran");

                //조출(휴일)
                oForm.DataSources.UserDataSources.Add("SEarlyTo", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("SEarlyTo").Specific.DataBind.SetBound(true, "", "SEarlyTo");

                //퇴근
                oForm.DataSources.UserDataSources.Add("OffDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("OffDate").Specific.DataBind.SetBound(true, "", "OffDate");

                oForm.DataSources.UserDataSources.Add("OffTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("OffTime").Specific.DataBind.SetBound(true, "", "OffTime");

                oForm.DataSources.UserDataSources.Add("Day2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Day2").Specific.DataBind.SetBound(true, "", "Day2");

                oForm.DataSources.UserDataSources.Add("DayOff2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("DayOff2").Specific.DataBind.SetBound(true, "", "DayOff2");

                //지각
                oForm.DataSources.UserDataSources.Add("LateTo", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("LateTo").Specific.DataBind.SetBound(true, "", "LateTo");

                //조퇴
                oForm.DataSources.UserDataSources.Add("EarlyOff", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("EarlyOff").Specific.DataBind.SetBound(true, "", "EarlyOff");

                //외출
                oForm.DataSources.UserDataSources.Add("GoOut", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("GoOut").Specific.DataBind.SetBound(true, "", "GoOut");

                //외출(시간)
                oForm.DataSources.UserDataSources.Add("GOFrTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("GOFrTime").Specific.DataBind.SetBound(true, "", "GOFrTime");

                oForm.DataSources.UserDataSources.Add("GOToTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("GOToTime").Specific.DataBind.SetBound(true, "", "GOToTime");

                oForm.DataSources.UserDataSources.Add("GOFrTim2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("GOFrTim2").Specific.DataBind.SetBound(true, "", "GOFrTim2");

                oForm.DataSources.UserDataSources.Add("GOToTim2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("GOToTim2").Specific.DataBind.SetBound(true, "", "GOToTim2");

                //직무코드
                oForm.DataSources.UserDataSources.Add("ActCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ActCode").Specific.DataBind.SetBound(true, "", "ActCode");

                //직무코드명
                oForm.DataSources.UserDataSources.Add("ActText", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ActText").Specific.DataBind.SetBound(true, "", "ActText");

                //위해
                oForm.DataSources.UserDataSources.Add("DangerCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("DangerCD").Specific.DataBind.SetBound(true, "", "DangerCD");

                //기찰기 출근
                oForm.DataSources.UserDataSources.Add("RToDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("RToDate").Specific.DataBind.SetBound(true, "", "RToDate");

                oForm.DataSources.UserDataSources.Add("RToTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("RToTime").Specific.DataBind.SetBound(true, "", "RToTime");

                //기찰기 퇴근
                oForm.DataSources.UserDataSources.Add("ROffDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ROffDate").Specific.DataBind.SetBound(true, "", "ROffDate");

                oForm.DataSources.UserDataSources.Add("ROffTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("ROffTime").Specific.DataBind.SetBound(true, "", "ROffTime");

                //기찰기 외출
                oForm.DataSources.UserDataSources.Add("ROutTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("ROutTime").Specific.DataBind.SetBound(true, "", "ROutTime");

                oForm.DataSources.UserDataSources.Add("RInTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("RInTime").Specific.DataBind.SetBound(true, "", "RInTime");

                //근태기찰정상확인
                oForm.DataSources.UserDataSources.Add("Confirm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("Confirm").Specific.DataBind.SetBound(true, "", "Confirm");

                //Check 카운터
                oForm.DataSources.UserDataSources.Add("Chkcnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("Chkcnt").Specific.DataBind.SetBound(true, "", "Chkcnt");
                //////////저장//////////_E

                //담당
                oForm.Items.Item("SRspCode").DisplayDesc = true;

                //근무형태
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN= 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ShiftDat").Specific, "Y");
                oForm.Items.Item("ShiftDat").DisplayDesc = true;

                //근무조
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "Y");
                oForm.Items.Item("GNMUJO").DisplayDesc = true;

                //근태
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P221' AND U_UseYN= 'Y' Order by U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("WorkType").Specific, "Y");
                oForm.Items.Item("WorkType").DisplayDesc = true;

                //위해
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P220' And U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "' AND U_UseYN= 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("DangerCD").Specific, "Y");
                oForm.Items.Item("DangerCD").DisplayDesc = true;
                
                //휴일
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("SDayOff").Specific, "N");
                oForm.Items.Item("SDayOff").DisplayDesc = true;

                //휴일
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("DayOff1").Specific, "N");
                oForm.Items.Item("DayOff1").DisplayDesc = true;

                //휴일
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("DayOff2").Specific, "N");
                oForm.Items.Item("DayOff2").DisplayDesc = true;

                //근태이상자 확인
                oForm.Items.Item("Confirm").Specific.ValidValues.Add("N", "미확인[N]");
                oForm.Items.Item("Confirm").Specific.ValidValues.Add("Y", "확인[Y]");
                oForm.Items.Item("Confirm").DisplayDesc = true;

                //교대근무인정
                oForm.Items.Item("RotateYN").Specific.ValidValues.Add("Y", "인정[Y]");
                oForm.Items.Item("RotateYN").Specific.ValidValues.Add("N", "미인정[N]");
                oForm.Items.Item("RotateYN").DisplayDesc = true;

                // oForm.Items("SPosDate").Specific.Value = Format(Now, "yyyymmdd")

                //1210 창원관리담당, 2210 부산 관리담당, 3210 포장사업장 지원담당
                sQry = "  Select  TeamCode = U_TeamCode,";
                sQry += "         RspCode = Isnull(U_RspCode,'')";
                sQry += " From    [@PH_PY001A]";
                sQry += " Where   Code = '" + dataHelpClass.User_MSTCOD() + "'";
                oRecordSet.DoQuery(sQry);

                //관리담당이면 현재일자 이전및 당일 16시 이후에 수정삭제 가능. 그외는 제한을 둠
                if (codeHelpClass.Right(oRecordSet.Fields.Item(0).Value.ToString().Trim(), 3) == "200" && oRecordSet.Fields.Item(1).Value.ToString().Trim() == "") // 운영지원팀
                {
                    oRspCodeYN = "Y";
                }
                else if (codeHelpClass.Right(oRecordSet.Fields.Item(1).Value.ToString().Trim(), 3) == "210") // 관리담당
                {
                    oRspCodeYN = "Y";
                }
                else
                {
                    oRspCodeYN = "N";
                }
                
                oForm.Update();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 세팅(Enable)
        /// </summary>
        private void PH_PY008_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", false); //행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면(Form) 아이템 세팅(Enable)
        /// </summary>
        private void PH_PY008_FormItemEnabled()
        {
            string sQry;
            int i;
            string CLTCOD;
            string sPosDate;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    CLTCOD = dataHelpClass.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'", "");
                    oForm.Items.Item("SCLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);

                    oForm.Items.Item("MSTCOD").Enabled = true;
                    oForm.Items.Item("FullName").Enabled = false;
                    oForm.Items.Item("TeamCode").Enabled = true;
                    oForm.Items.Item("RspCode").Enabled = true;
                    oForm.Items.Item("SCLTCOD").Enabled = true;

                    oForm.Items.Item("SPosDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

                    oForm.Items.Item("MSTCOD").Specific.Value = "";
                    oForm.Items.Item("FullName").Specific.Value = "";

                    if (oForm.Items.Item("STeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("STeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("STeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("STeamCode").Specific.ValidValues.Add("", "");

                        sQry = "  SELECT      U_Code,";
                        sQry += "             U_CodeNm";
                        sQry += " FROM        [@PS_HR200L] ";
                        sQry += " WHERE       Code = '1'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    U_Seq";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("STeamCode").Specific, "Y");
                    }
                    oForm.Items.Item("STeamCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //부서
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");

                        sQry = "  SELECT      U_Code,";
                        sQry += "             U_CodeNm";
                        sQry += " FROM        [@PS_HR200L] ";
                        sQry += " WHERE       Code = '1'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += " ORDER BY    U_Seq";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                    }
                    oForm.Items.Item("TeamCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //담당
                    oForm.Items.Item("RspCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //반
                    oForm.Items.Item("ClsCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    
                    //근무형태
                    oForm.Items.Item("ShiftDat").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //근무조
                    oForm.Items.Item("GNMUJO").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    oForm.Items.Item("SShiftDat").Specific.Value = "";
                    oForm.Items.Item("ShiftDatNm").Specific.Value = "";
                    oForm.Items.Item("SGNMUJO").Specific.Value = "";
                    oForm.Items.Item("SGNMUJONm").Specific.Value = "";

                    oForm.Items.Item("ActCode").Specific.Value = "";
                    oForm.Items.Item("ActText").Specific.Value = "";

                    oForm.DataSources.UserDataSources.Item("RToTime").Value = "0000";
                    oForm.DataSources.UserDataSources.Item("ROffTime").Value = "0000";
                    oForm.DataSources.UserDataSources.Item("ROutTime").Value = "0000";
                    oForm.DataSources.UserDataSources.Item("RInTime").Value = "0000";

                    oForm.DataSources.UserDataSources.Item("RToDate").Value = "";
                    oForm.DataSources.UserDataSources.Item("ROffDate").Value = "";

                    oForm.DataSources.UserDataSources.Item("Comment").Value = ""; //비고 추가(2013.11.19 송명규)

                    CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                    sPosDate = oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim();

                    sQry = "  SELECT      U_WorkType";
                    sQry += " FROM        [@PH_PY003A] AS a";
                    sQry += "             Inner Join";
                    sQry += "             [@PH_PY003B] AS b";
                    sQry += "                 On a.Code = b.Code";
                    sQry += " WHERE       a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += "             AND b.U_Date = '" + sPosDate + "'";
                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount == 0)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("근태월력의 근무일 등록을 하지않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        oForm.Items.Item("WorkType").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else
                    {
                        oForm.Items.Item("WorkType").Specific.Select(PH_PY008_WorkTypeSelect(oForm.Items.Item("SPosDate").Specific.Value), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("FullName").Enabled = true;
                    oForm.Items.Item("CLTCOD").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "SCLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("FullName").Enabled = false;
                    oForm.Items.Item("CLTCOD").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 근태구분 조회
        /// </summary>
        /// <param name="pDate">일자</param>
        /// <returns>근태구분</returns>
        private string PH_PY008_WorkTypeSelect(string pDate)
        {
            string sQry;
            string returnValue = string.Empty;
            short errNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT U_WorkType From [@PH_PY003B] Where U_Date = '" + pDate + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    returnValue = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                }
                else
                {
                    errNum = 1;
                    throw new Exception();
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 그리드 타이틀 세팅
        /// </summary>
        /// <param name="iRow"></param>
        private void PH_PY008_TitleSetting(int iRow)
        {
            int i;
            int j;
            string sQry;
            string[] COLNAM = new string[30];

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "순번";
                COLNAM[1] = "선택";
                COLNAM[2] = "부서";
                COLNAM[3] = "담당";
                COLNAM[4] = "반";
                COLNAM[5] = "사번";
                COLNAM[6] = "성명";
                COLNAM[7] = "요일";
                COLNAM[8] = "휴일";
                COLNAM[9] = "기본";
                COLNAM[10] = "연장";
                COLNAM[11] = "심야";
                COLNAM[12] = "특근";
                COLNAM[13] = "특연";
                COLNAM[14] = "조출";
                COLNAM[15] = "휴조";
                COLNAM[16] = "교육";
                COLNAM[17] = "지각";
                COLNAM[18] = "조퇴";
                COLNAM[19] = "외출";
                COLNAM[20] = "위해";
                COLNAM[21] = "직무";
                COLNAM[22] = "근태";
                COLNAM[23] = "교대일수";
                COLNAM[24] = "출근일자";
                COLNAM[25] = "출근시간";
                COLNAM[26] = "퇴근일자";
                COLNAM[27] = "퇴근시간";
                COLNAM[28] = "비고";
                COLNAM[29] = "교대인정대상";

                for (i = 0; i < COLNAM.Length; i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    oGrid1.Columns.Item(i).Editable = false;

                    switch (COLNAM[i])
                    {
                        case "선택":
                            
                            oGrid1.Columns.Item(i).Editable = true;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                            break;
                            
                        case "부서":
                            
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            
                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '1' AND U_UseYN= 'Y' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("TeamCode")).ValidValues.Add("", "");
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("TeamCode")).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet.MoveNext();
                                }
                            }

                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("TeamCode")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                            
                        case "담당":
                            
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            
                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '2' AND U_UseYN= 'Y' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("RspCode")).ValidValues.Add("", "");
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("RspCode")).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet.MoveNext();
                                }
                            }

                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("RspCode")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                            

                        case "반":
                            
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = '9' AND U_UseYN= 'Y' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("ClsCode")).ValidValues.Add("", "");
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("ClsCode")).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet.MoveNext();
                                }
                            }

                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("ClsCode")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                            
                        case "휴일":
                            
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202'";
                            oRecordSet.DoQuery(sQry);
                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("DayOff")).ValidValues.Add("", "");
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("DayOff")).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet.MoveNext();
                                }
                            }

                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("DayOff")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                            
                        case "근태":
                            
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            
                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P221' AND U_UseYN= 'Y' Order by U_Seq";
                            oRecordSet.DoQuery(sQry);
                            if (oRecordSet.RecordCount > 0)
                            {
                                for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                                {
                                    ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("WorkType")).ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                    oRecordSet.MoveNext();
                                }
                            }

                            ((SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("WorkType")).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;
                            
                        case "기본":
                        case "연장":
                        case "심야":
                        case "특근":
                        case "특연":
                        case "조출":
                        case "휴조":
                        case "교육":
                        case "지각":
                        case "조퇴":
                        case "외출":
                            
                            oGrid1.Columns.Item(i).RightJustified = true;
                            break;
                            
                        default:
                            
                            oGrid1.Columns.Item(i).Editable = false;
                            break;
                    }
                }

                oGrid1.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); 
            }
        }
        
        /// <summary>
        /// 그리드 데이터 로드
        /// </summary>
        /// <param name="oConfirm">oConfirm A: 조회, S:근태이상자</param>
        private void PH_PY008_MTX01(string oConfirm)
        {
            int i;
            int Chkcnt = 0;
            int iRow;
            short errNum = 0;
            string sQry;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;

            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                oForm.Freeze(true);

                Param01 = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("STeamCode").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("SRspCode").Specific.Value.ToString().Trim();
                Param05 = oForm.Items.Item("SClsCode").Specific.Value.ToString().Trim();
                Param06 = oForm.Items.Item("SMSTCOD").Specific.Value.ToString().Trim();
                Param07 = oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim();
                Param08 = oForm.Items.Item("SGNMUJO").Specific.Value.ToString().Trim();

                //해당일의 근무일 등록 확인
                sQry = "  Select  U_WorkType";
                sQry += " From    [@PH_PY003A] a";
                sQry += "         Inner Join";
                sQry += "         [@PH_PY003B] b";
                sQry += "             On a.Code = b.Code ";
                sQry += " Where   a.U_CLTCOD = '" + Param01 + "'";
                sQry += "         And b.U_Date = '" + Param02 + "'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC PH_PY008_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + oConfirm + "'";
                oDS_PH_PY008.ExecuteQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                }

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY008_TitleSetting(iRow);
                PH_PY008_MTX02("", 0, "");

                if (oDS_PH_PY008.Rows.Count > 0)
                {
                    for (i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value.ToString().Trim() == "Y")
                        {
                            Chkcnt += 1;
                        }
                    }

                    oForm.Items.Item("Chkcnt").Specific.Value = Chkcnt;
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근무일 등록을 하지않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                ProgBar01.Value = 100;
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 그리드 데이터 로드(실제)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY008_MTX02(string oUID, int oRow, string oCol)
        {   
            int i;
            int sRow;
            short errNum = 0;
            string sQry;
            string Param01;
            string Param02;
            string Param03;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                sRow = oRow;

                Param01 = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                Param02 = oDS_PH_PY008.Columns.Item("GetDate").Cells.Item(oRow).Value == null ? "" : Convert.ToDateTime(oDS_PH_PY008.Columns.Item("GetDate").Cells.Item(oRow).Value.ToString().Trim()).ToString("yyyyMMdd"); //GetDate Null 처리 추가;
                Param03 = oDS_PH_PY008.Columns.Item("MSTCOD").Cells.Item(oRow).Value.ToString().Trim();

                if (Param02 == "") //"출근" 필드에 일자가 없으면
                {
                    oForm.Items.Item("TeamCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    oForm.Items.Item("MSTCOD").Specific.Value = "";
                    oForm.Items.Item("FullName").Specific.Value = "";

                    oForm.DataSources.UserDataSources.Item("GetTime").Value = "0000";
                    oForm.DataSources.UserDataSources.Item("OffTime").Value = "0000";

                    PH_PY008_Time_ReSet();

                    oForm.Items.Item("PosDate").Enabled = true;
                    oForm.Items.Item("TeamCode").Enabled = true;
                    oForm.Items.Item("RspCode").Enabled = true;
                    oForm.Items.Item("ClsCode").Enabled = true;
                    oForm.Items.Item("MSTCOD").Enabled = true;
                    oForm.Items.Item("FullName").Enabled = true;

                    errNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC PH_PY008_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 2;
                    throw new Exception();
                }

                oForm.DataSources.UserDataSources.Item("PosDate").Value = oRecordSet.Fields.Item("PosDate").Value.ToString("yyyyMMdd");

                oForm.Items.Item("SDayOff").Specific.Select(oRecordSet.Fields.Item("DayOff").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.DataSources.UserDataSources.Item("SDay").Value = oRecordSet.Fields.Item("Day").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("FullName").Value = oRecordSet.Fields.Item("FullName").Value.ToString().Trim();

                //부서
                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                {
                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                }

                sQry = "  SELECT      U_Code,";
                sQry += "             U_CodeNm";
                sQry += " FROM        [@PS_HR200L]";
                sQry += " WHERE       Code = '1' ";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                oForm.Items.Item("TeamCode").Specific.Select(oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                
                //담당
                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                {
                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                }

                sQry = "  SELECT      U_Code,";
                sQry += "             U_CodeNm";
                sQry += " FROM        [@PS_HR200L] ";
                sQry += " WHERE       Code = '2' ";
                sQry += " ORDER BY    U_Seq";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                oForm.Items.Item("RspCode").Specific.Select(oRecordSet.Fields.Item("RspCode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //반
                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                {
                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    oForm.Items.Item("ClsCode").Specific.ValidValues.Add("", "");
                }

                sQry = "  SELECT      U_Code,";
                sQry += "             U_CodeNm";
                sQry += " FROM        [@PS_HR200L] ";
                sQry += " WHERE       Code = '9' ";
                sQry += " Order By U_Seq";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");

                oForm.Items.Item("ClsCode").Specific.Select(oRecordSet.Fields.Item("ClsCode").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //근무형태
                oForm.Items.Item("ShiftDat").Specific.Select(oRecordSet.Fields.Item("ShiftDat").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //근무조
                if (oForm.Items.Item("GNMUJO").Specific.ValidValues.Count > 0)
                {
                    for (i = oForm.Items.Item("GNMUJO").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                    {
                        oForm.Items.Item("GNMUJO").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                        
                    oForm.Items.Item("GNMUJO").Specific.ValidValues.Add("", "");
                }

                sQry = "  SELECT      U_Code,";
                sQry += "             U_CodeNm";
                sQry += " FROM        [@PS_HR200L] ";
                sQry += " WHERE       Code = 'P155'";
                sQry += "             AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.Value + "'";
                sQry += "             And U_UseYN = 'Y'";
                sQry += " ORDER BY    U_Code";

                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "Y");

                oForm.Items.Item("GNMUJO").Specific.Select(oRecordSet.Fields.Item("GNMUJO").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("WorkType").Specific.Select(oRecordSet.Fields.Item("WorkType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue); //근태구분
                oForm.DataSources.UserDataSources.Item("GetDate").Value = oRecordSet.Fields.Item("GetDate").Value.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("GetTime").Value = codeHelpClass.Right("0" + oRecordSet.Fields.Item("GetTime").Value.ToString().Trim(), 4);
                oForm.DataSources.UserDataSources.Item("OffDate").Value = oRecordSet.Fields.Item("OffDate").Value.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("OffTime").Value = codeHelpClass.Right("0" + oRecordSet.Fields.Item("OffTime").Value.ToString().Trim(), 4);
                oForm.DataSources.UserDataSources.Item("Base").Value = oRecordSet.Fields.Item("Base").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("Extend").Value = oRecordSet.Fields.Item("Extend").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("Special").Value = oRecordSet.Fields.Item("Special").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("SpExtend").Value = oRecordSet.Fields.Item("SpExtend").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("Midnight").Value = oRecordSet.Fields.Item("Midnight").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("LateTo").Value = oRecordSet.Fields.Item("LateTo").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("EarlyOff").Value = oRecordSet.Fields.Item("EarlyOff").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("EarlyTo").Value = oRecordSet.Fields.Item("EarlyTo").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("SEarlyTo").Value = oRecordSet.Fields.Item("SEarlyTo").Value.ToString().Trim(); // 휴일조출(2013.06.17 송명규 추가)
                oForm.DataSources.UserDataSources.Item("EducTran").Value = oRecordSet.Fields.Item("EducTran").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("GoOut").Value = oRecordSet.Fields.Item("GoOut").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("GOFrTime").Value = oRecordSet.Fields.Item("GOFrTime").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("GOFrTim2").Value = oRecordSet.Fields.Item("GOFrTim2").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("GOToTime").Value = oRecordSet.Fields.Item("GOToTime").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("GOToTim2").Value = oRecordSet.Fields.Item("GOToTim2").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("ActCode").Value = oRecordSet.Fields.Item("ActCode").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("ActText").Value = oRecordSet.Fields.Item("ActText").Value.ToString().Trim();
                oForm.Items.Item("DangerCD").Specific.Select(oRecordSet.Fields.Item("DangerCD").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("Rotation").Value = oRecordSet.Fields.Item("Rotation").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("RToDate").Value = oRecordSet.Fields.Item("RToDate").Value.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("RToTime").Value = oRecordSet.Fields.Item("RToTime").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("ROffDate").Value = oRecordSet.Fields.Item("ROffDate").Value.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("ROffTime").Value = oRecordSet.Fields.Item("ROffTime").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("ROutTime").Value = oRecordSet.Fields.Item("ROutTime").Value.ToString().Trim();
                oForm.DataSources.UserDataSources.Item("RInTime").Value = oRecordSet.Fields.Item("RInTime").Value.ToString().Trim();
                oForm.Items.Item("Confirm").Specific.Select(oRecordSet.Fields.Item("Confirm").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("RotateYN").Specific.Select(oRecordSet.Fields.Item("RotateYN").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.DataSources.UserDataSources.Item("Comment").Value = oRecordSet.Fields.Item("Comment").Value.ToString().Trim(); //비고(2013.11.19 송명규 추가)

                oForm.Items.Item("PosDate").Enabled = false;
                oForm.Items.Item("TeamCode").Enabled = false;
                oForm.Items.Item("RspCode").Enabled = false;
                oForm.Items.Item("ClsCode").Enabled = false;
                oForm.Items.Item("MSTCOD").Enabled = false;
                oForm.Items.Item("FullName").Enabled = false;

                oForm.Update();
            }
            catch(Exception ex)
            {
                if(errNum == 1)
                {
                    //오류메시지 없이 메소드 종료
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 시간초기화
        /// </summary>
        private void PH_PY008_Time_ReSet()
        {
            try
            {
                oForm.Freeze(true);
                oForm.DataSources.UserDataSources.Item("Base").Value = "0";
                oForm.DataSources.UserDataSources.Item("Extend").Value = "0";
                oForm.DataSources.UserDataSources.Item("Midnight").Value = "0";
                oForm.DataSources.UserDataSources.Item("EarlyTo").Value = "0";
                oForm.DataSources.UserDataSources.Item("Special").Value = "0";
                oForm.DataSources.UserDataSources.Item("SpExtend").Value = "0";
                oForm.DataSources.UserDataSources.Item("EducTran").Value = "0";
                oForm.DataSources.UserDataSources.Item("SEarlyTo").Value = "0";
                oForm.DataSources.UserDataSources.Item("LateTo").Value = "0";
                oForm.DataSources.UserDataSources.Item("EarlyOff").Value = "0";
                oForm.DataSources.UserDataSources.Item("GoOut").Value = "0";

                oForm.DataSources.UserDataSources.Item("GOFrTime").Value = "0000";
                oForm.DataSources.UserDataSources.Item("GOToTime").Value = "0000";
                oForm.DataSources.UserDataSources.Item("GOFrTim2").Value = "0000";
                oForm.DataSources.UserDataSources.Item("GOToTim2").Value = "0000";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 한글 요일 반환
        /// </summary>
        /// <param name="pDate">일자</param>
        /// <returns>요일(한글)</returns>
        private string PH_PY008_DaySelect(string pDate)
        {
            string returnValue = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (Convert.ToDateTime(dataHelpClass.ConvertDateType(pDate, "-")).DayOfWeek)
                {

                    case DayOfWeek.Monday:
                        returnValue = "월";
                        break;
                    case DayOfWeek.Tuesday:
                        returnValue = "화";
                        break;
                    case DayOfWeek.Wednesday:
                        returnValue = "수";
                        break;
                    case DayOfWeek.Thursday:
                        returnValue = "목";
                        break;
                    case DayOfWeek.Friday:
                        returnValue = "금";
                        break;
                    case DayOfWeek.Saturday:
                        returnValue = "토";
                        break;
                    case DayOfWeek.Sunday:
                        returnValue = "일";
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return returnValue;
        }

        /// <summary>
        /// 평일, 휴일 반환
        /// </summary>
        /// <param name="pDate">일자</param>
        /// <returns>일자에 따른 평일, 휴일</returns>
        private string PH_PY008_DayOffSelect(string pDate)
        {
            string sQry;
            string CLTCOD; // 사업장
            string retunValue = string.Empty;
            short errNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;

                sQry = "  SELECT      T1.U_DayType ";
                sQry += " FROM        [@PH_PY003A] AS T0";
                sQry += "             INNER JOIN";
                sQry += "             [@PH_PY003B] AS T1";
                sQry += "                 ON T0.Code = T1.Code";
                sQry += " WHERE       T1.U_Date = '" + pDate + "'";
                sQry += "             AND T0.U_CLTCOD = '" + CLTCOD + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    retunValue = oRecordSet.Fields.Item(0).Value;
                }
                else
                {
                    errNum = 1;
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return retunValue;
        }

        /// <summary>
        /// 데이터 저장
        /// </summary>
        private void PH_PY008_SAVE()
        {
            int i;
            string sQry;
            short errNum = 0;

            string GNMUJO;
            string ClsCode;
            string RspCode;
            string FullName;
            string pDate;
            string CLTCOD;
            string DayOff;
            string MSTCOD;
            string TeamCode;
            string Inform05;
            string ShiftDat;
            string WorkType;
            string GetDate;
            short GetTime;
            string OffDate;
            short OffTime;
            double Base = 0;
            double Extend = 0;
            double Special = 0;
            double SpExtend = 0;
            double Midnight;
            double LateTo;
            double EarlyOff;
            double EarlyTo = 0;
            double SEarlyTo = 0;
            double EducTran;
            double GoOut;
            short GOFrTime;
            short GOToTime;
            short GOFrTim2;
            short GoToTim2;
            string ActText;
            string ActCode;
            string DangerCD;
            double DangerNu;
            double Rotation;
            string RToDate;
            short RToTime;
            string ROffDate;
            short ROffTime;
            short ROutTime;
            short RInTime;
            string Confirm;
            string Abnomal;
            string Attend = string.Empty;
            string PosDate;
            string DAYTyp;
            string ShiftDat001A; //인사기본 교대조
            string bShiftDat; //지난주 교대조
            string RotateYN;
            string CheckData;
            double OverTime = 0;
            string OverTimeYN;
            double SpecialTot; //사무기술직 특근24시간 제한
            string JIGTYP; //직원구분02 : 사무기술직
            string JIGCOD; //직급(2014.07.23 송명규 추가)
            string PAYTYP; //급여형태 2:월급제
            string Comment; //비고(2013.11.19 송명규 추가)
            string BPLID2; //근무지사업장(2018.10.01 송명규 추가)

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                RotateYN = "N"; //야간교대조 인정 초기값 N

                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                PosDate = oForm.Items.Item("PosDate").Specific.Value.ToString().Trim();
                DayOff = oForm.Items.Item("SDayOff").Specific.Value.ToString().Trim();
                pDate = oForm.Items.Item("SDay").Specific.Value.ToString().Trim();
                ShiftDat = oForm.Items.Item("ShiftDat").Specific.Value.ToString().Trim();
                GNMUJO = oForm.Items.Item("GNMUJO").Specific.Value.ToString().Trim();
                WorkType = oForm.Items.Item("WorkType").Specific.Value.ToString().Trim();
                GetDate = oForm.Items.Item("GetDate").Specific.Value.ToString().Trim();
                GetTime = Convert.ToInt16(oForm.Items.Item("GetTime").Specific.Value.ToString().Trim());
                OffDate = oForm.Items.Item("OffDate").Specific.Value.ToString().Trim();
                OffTime = Convert.ToInt16(oForm.Items.Item("OffTime").Specific.Value.ToString().Trim());
                Base = Convert.ToSingle(oForm.Items.Item("Base").Specific.Value.ToString().Trim());
                Extend = Convert.ToSingle(oForm.Items.Item("Extend").Specific.Value.ToString().Trim());
                Special = Convert.ToSingle(oForm.Items.Item("Special").Specific.Value.ToString().Trim());
                SpExtend = Convert.ToSingle(oForm.Items.Item("SpExtend").Specific.Value.ToString().Trim());
                Midnight = Convert.ToSingle(oForm.Items.Item("Midnight").Specific.Value.ToString().Trim());
                LateTo = Convert.ToSingle(oForm.Items.Item("LateTo").Specific.Value.ToString().Trim());
                EarlyOff = Convert.ToSingle(oForm.Items.Item("EarlyOff").Specific.Value.ToString().Trim());
                EarlyTo = Convert.ToSingle(oForm.Items.Item("EarlyTo").Specific.Value.ToString().Trim());
                SEarlyTo = Convert.ToSingle(oForm.Items.Item("SEarlyTo").Specific.Value.ToString().Trim());
                EducTran = Convert.ToSingle(oForm.Items.Item("EducTran").Specific.Value.ToString().Trim());
                GoOut = Convert.ToSingle(oForm.Items.Item("GoOut").Specific.Value.ToString().Trim());
                GOFrTime = Convert.ToInt16(oForm.Items.Item("GOFrTime").Specific.Value.ToString().Trim());
                GOToTime = Convert.ToInt16(oForm.Items.Item("GOToTime").Specific.Value.ToString().Trim());
                GOFrTim2 = Convert.ToInt16(oForm.Items.Item("GOFrTim2").Specific.Value.ToString().Trim());
                GoToTim2 = Convert.ToInt16(oForm.Items.Item("GOToTim2").Specific.Value.ToString().Trim());
                RToDate = oForm.Items.Item("RToDate").Specific.Value.ToString().Trim();
                RToTime = Convert.ToInt16(oForm.Items.Item("RToTime").Specific.Value.ToString().Trim());
                ROffDate = oForm.Items.Item("ROffDate").Specific.Value.ToString().Trim();
                ROffTime = Convert.ToInt16(oForm.Items.Item("ROffTime").Specific.Value.ToString().Trim());
                ROutTime = Convert.ToInt16(oForm.Items.Item("ROutTime").Specific.Value.ToString().Trim());
                RInTime = Convert.ToInt16(oForm.Items.Item("RInTime").Specific.Value.ToString().Trim());

                Confirm = oForm.Items.Item("Confirm").Specific.Value.ToString().Trim();
                Comment = oForm.Items.Item("Comment").Specific.Value.ToString().Trim();

                if (GNMUJO == "41")
                {
                    DAYTyp = oForm.Items.Item("Team1").Specific.Value.ToString().Trim();
                }
                else if (GNMUJO == "42")
                {
                    DAYTyp = oForm.Items.Item("Team2").Specific.Value.ToString().Trim();
                }
                else
                {
                    DAYTyp = oForm.Items.Item("Team3").Specific.Value.ToString().Trim();
                }

                if (Confirm == "")
                {
                    Confirm = "N";
                }
                
                if (GNMUJO == "")
                {
                    errNum = 11;
                    throw new Exception();
                }

                if (oForm.DataSources.UserDataSources.Item("Chk").Value.ToString().Trim() == "N")
                {
                    //전체적용
                    if (oDS_PH_PY008.Rows.Count > 0)
                    {
                        for (i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                        {
                            if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value.ToString().Trim() == "Y")
                            {
                                CheckData = oDS_PH_PY008.Columns.Item("Day").Cells.Item(i).Value.ToString().Trim();
                                MSTCOD = oDS_PH_PY008.Columns.Item("MSTCOD").Cells.Item(i).Value.ToString().Trim();
                                FullName = oDS_PH_PY008.Columns.Item("FullName").Cells.Item(i).Value.ToString().Trim();
                                BPLID2 = dataHelpClass.Get_ReData("U_BPLID2", "Code", "[@PH_PY001A]", "'" + MSTCOD + "'", ""); //근무지사업장(2018.10.01 송명규)
                                bShiftDat = oDS_PH_PY008.Columns.Item("bShiftDat").Cells.Item(i).Value.ToString().Trim(); //지난주 교대조

                                //사원의 직무와 위해 조회
                                sQry = "  select  Activity = a.U_Activity, ";
                                sQry += "         ActText = b.CodeNm, ";
                                sQry += "         DangerCD = b.DangerCD, ";
                                sQry += "         ShiftDat = a.U_ShiftDat, ";
                                sQry += "         JIGTYP = a.U_JIGTYP, ";
                                sQry += "         JIGCOD = a.U_JIGCOD, "; //직급(2014.07.23 송명규 추가)
                                sQry += "         PAYTYP = a.U_PAYTYP, ";
                                sQry += "         TeamCode = a.U_TeamCode, ";
                                sQry += "         RspCode = a.U_RspCode, ";
                                sQry += "         ClsCode = a.U_ClsCode, ";
                                sQry += "         inform05 = a.u_inform05 ";
                                sQry += " FROM    [@PH_PY001A] AS a";
                                sQry += "         Left Join";
                                sQry += "         (";
                                sQry += "             Select  Code = t1.U_Code,";
                                sQry += "                     CodeNm = t1.U_CodeNm,";
                                sQry += "                     DangerCD = t1.U_Char1 ";
                                sQry += "             FROM    [@PS_HR200H] t";
                                sQry += "                     Inner Join";
                                sQry += "                     [@PS_HR200L] t1";
                                sQry += "                         On t.Code = t1.Code ";
                                sQry += "             Where   t1.Code = 'P127'";
                                sQry += "                     and U_Char2 = '" + CLTCOD + "'";
                                sQry += "         ) AS b";
                                sQry += "             On a.U_Activity = b.Code";
                                sQry += " Where   a.Code = '" + MSTCOD + "'";

                                oRecordSet.DoQuery(sQry);

                                ActCode = oRecordSet.Fields.Item("Activity").Value.ToString().Trim();
                                ActText = oRecordSet.Fields.Item("ActText").Value.ToString().Trim();
                                DangerCD = oRecordSet.Fields.Item("DangerCD").Value.ToString().Trim();

                                JIGTYP = oRecordSet.Fields.Item("JIGTYP").Value.ToString().Trim(); //직원구분
                                JIGCOD = oRecordSet.Fields.Item("JIGCOD").Value.ToString().Trim(); //직급
                                PAYTYP = oRecordSet.Fields.Item("PAYTYP").Value.ToString().Trim(); //급여구분

                                TeamCode = oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim(); //팀
                                RspCode = oRecordSet.Fields.Item("RspCode").Value.ToString().Trim(); //담당
                                ClsCode = oRecordSet.Fields.Item("ClsCode").Value.ToString().Trim(); //반
                                Inform05 = oRecordSet.Fields.Item("inform05").Value.ToString().Trim();
                                ShiftDat001A = oRecordSet.Fields.Item("ShiftDat").Value.ToString().Trim(); //인사기본 근무조

                                //1교대자가 2교대야간 근무를 근무했을시 1주일 후 같은 요일에 야간근무인정
                                //하루 야간근무시 2일 야근교대수당을 지급 2016/01/28 NGY추가
                                if (ShiftDat001A == "1")
                                {
                                    if (bShiftDat == "2")
                                    {
                                        RotateYN = "Y";
                                    }
                                    else
                                    {
                                        RotateYN = "N";
                                    }
                                }
                                
                                if (CheckData == "") //근태구분이 없으면 전체입력대상
                                {
                                    Abnomal = "Y"; //근태이상자

                                    if (oForm.Items.Item("SDay").Specific.Value == "월") //월요일은 제외
                                    {
                                        OverTime = 0;
                                    }
                                    else
                                    {
                                        OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD);
                                    }
                                    
                                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime > 52) //스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                                    {
                                        if (Inform05 == "Y")
                                        {
                                            if (PSH_Globals.SBO_Application.MessageBox("주 52시간 연장초과 했습니다. 그래도 입력하시겠습니까?", 2, "예", "아니오") == 1)
                                            {
                                                OverTime = 0;
                                            }
                                        }
                                        else
                                        {
                                            errNum = 1;
                                            throw new Exception();
                                        }
                                    }

                                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52)
                                    {
                                        if (LateTo + EarlyOff + GoOut > 0.1) //지각, 조퇴, 외출시간이 있으면 교대일수 없다.
                                        {
                                            Rotation = 0;
                                        }
                                        else
                                        {
                                            Rotation = 1;
                                        }

                                        if (codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11")
                                        {
                                            Rotation = 0;
                                        }

                                        if ((JIGTYP == "02" || JIGTYP == "03" || JIGCOD == "73") && PAYTYP == "2")
                                        {
                                            if (CLTCOD != "2")
                                            {
                                                Extend = PH_PY008_Extend15(MSTCOD);

                                                SpecialTot = PH_PY008_Special(MSTCOD);

                                                if (SpecialTot >= 24)
                                                {
                                                    Special = 0;
                                                }
                                                else if ((Special + SpecialTot) >= 24)
                                                {
                                                    Special -= (Special + SpecialTot - 24);
                                                }
                                            }
                                            else
                                            {
                                                Extend = System.Convert.ToDouble(oForm.Items.Item("Extend").Specific.Value.ToString().Trim());
                                            }
                                        }
                                        else
                                        {
                                            Extend = System.Convert.ToDouble(oForm.Items.Item("Extend").Specific.Value.ToString().Trim());
                                        }
                                        
                                        if (WorkType == "A01" || WorkType == "A02" || codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11") //무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                                        {
                                            DangerCD = "";
                                        }
                                        else
                                        {
                                            
                                            if ((JIGTYP == "04" || JIGTYP == "05") && PAYTYP != "1" & JIGCOD != "73" && DangerCD == "") //전문직, 계약직 이며 연봉제가 아니고 위해코드가 없으면 위해코드를 기타로..
                                            {
                                                if (CLTCOD == "1")
                                                {
                                                    DangerCD = "31";
                                                }
                                            }
                                        }

                                        if (Base + Special > 0) //근무시간이 있으나 4시간미만이면 위해수당 없다.
                                        {
                                            if (Special + SpExtend + Base + Extend < 4)
                                            {
                                                DangerCD = "";
                                            }
                                        }
                                        else //근무시간이 없으면 위해코드를 기타로.
                                        {
                                            if (CLTCOD == "1")
                                            {
                                                if (DangerCD.Trim() != "")
                                                {
                                                    DangerCD = "31";
                                                }
                                            }
                                        }

                                        if (DangerCD == "")
                                        {
                                            DangerNu = 0;
                                        }
                                        else
                                        {
                                            DangerNu = 1;
                                        }

                                        if (CLTCOD == "3")
                                        {
                                            DangerCD = "";
                                            DangerNu = 0;
                                        }

                                        sQry = "INSERT INTO ZPH_PY008";
                                        sQry += " (";
                                        sQry += " CLTCOD,";
                                        sQry += " PosDate,";
                                        sQry += " DayOff,";
                                        sQry += " Day,";
                                        sQry += " MSTCOD,";
                                        sQry += " FullName,";
                                        sQry += " TeamCode,";
                                        sQry += " RspCode,";
                                        sQry += " ClsCode,";
                                        sQry += " ShiftDat,";
                                        sQry += " GNMUJO,";
                                        sQry += " WorkType,";
                                        sQry += " GetDate,";
                                        sQry += " GetTime,";
                                        sQry += " OffDate,";
                                        sQry += " OffTime,";
                                        sQry += " Base,";
                                        sQry += " Extend,";
                                        sQry += " Special,";
                                        sQry += " SpExtend,";
                                        sQry += " Midnight,";
                                        sQry += " LateTo,";
                                        sQry += " EarlyOff,";
                                        sQry += " EarlyTo,";
                                        sQry += " SEarlyTo,";
                                        sQry += " EducTran,";
                                        sQry += " GoOut,";
                                        sQry += " GOFrTime,";
                                        sQry += " GOToTime,";
                                        sQry += " GOFrTim2,";
                                        sQry += " GOToTim2,";
                                        sQry += " ActCode,";
                                        sQry += " ActText,";
                                        sQry += " DangerCD,";
                                        sQry += " DangerNu,";
                                        sQry += " Rotation,";
                                        sQry += " RToDate,";
                                        sQry += " RToTime,";
                                        sQry += " ROffDate,";
                                        sQry += " ROffTime,";
                                        sQry += " ROutTime,";
                                        sQry += " RInTime,";
                                        sQry += " Abnomal,";
                                        sQry += " Attend,";
                                        sQry += " Confirm,";
                                        sQry += " Comment,";
                                        sQry += " RotateYN,";
                                        sQry += " DAYTyp,";
                                        sQry += " BPLID2";
                                        sQry += " ) ";
                                        sQry += "VALUES(";
                                        sQry += "'" + CLTCOD + "',";
                                        sQry += "'" + PosDate + "',";
                                        sQry += "'" + DayOff + "',";
                                        sQry += "'" + pDate + "',";
                                        sQry += "'" + MSTCOD + "',";
                                        sQry += "'" + FullName + "',";
                                        sQry += "'" + TeamCode + "',";
                                        sQry += "'" + RspCode + "',";
                                        sQry += "'" + ClsCode + "',";
                                        sQry += "'" + ShiftDat + "',";
                                        sQry += "'" + GNMUJO + "',";
                                        sQry += "'" + WorkType + "',";
                                        sQry += "'" + GetDate + "',";
                                        sQry += GetTime + ",";
                                        sQry += "'" + OffDate + "',";
                                        sQry += OffTime + ",";
                                        sQry += Base + ",";
                                        sQry += Extend + ",";
                                        sQry += Special + ",";
                                        sQry += SpExtend + ",";
                                        sQry += Midnight + ",";
                                        sQry += LateTo + ",";
                                        sQry += EarlyOff + ",";
                                        sQry += EarlyTo + ",";
                                        sQry += SEarlyTo + ",";
                                        sQry += EducTran + ",";
                                        sQry += GoOut + ",";
                                        sQry += GOFrTime + ",";
                                        sQry += GOToTime + ",";
                                        sQry += GOFrTim2 + ",";
                                        sQry += GoToTim2 + ",";
                                        sQry += "'" + ActCode + "',";
                                        sQry += "'" + ActText + "',";
                                        sQry += "'" + DangerCD + "',";
                                        sQry += DangerNu + ",";
                                        sQry += Rotation + ",";
                                        sQry += "'" + RToDate + "',";
                                        sQry += RToTime + ",";
                                        sQry += "'" + ROffDate + "',";
                                        sQry += ROffTime + ",";
                                        sQry += ROutTime + ",";
                                        sQry += RInTime + ",";
                                        sQry += "'" + Abnomal + "',";
                                        sQry += "'" + Attend + "',";
                                        sQry += "'" + Confirm + "',";
                                        sQry += "'" + Comment + "',";
                                        sQry += "'" + RotateYN + "',";
                                        sQry += "'" + DAYTyp + "',";
                                        sQry += "'" + BPLID2 + "'";
                                        sQry += ")";
                                        oRecordSet.DoQuery(sQry);
                                    }
                                }
                                else
                                {
                                    if (oForm.Items.Item("SDay").Specific.Value.ToString().Trim() == "월") //월요일은 제외
                                    {
                                        sQry = "SELECT Isnull(U_Inform05,'N') From [@PH_PY001A] Where Code = '" + MSTCOD + "'";
                                        oRecordSet.DoQuery(sQry);

                                        OverTimeYN = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                                        if (OverTimeYN == "Y")
                                        {
                                            OverTime = -100;
                                        }
                                        else
                                        {
                                            OverTime = 0;
                                        }
                                    }
                                    else
                                    {
                                        OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD);
                                    }
                                    
                                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime > 52) //스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                                    {
                                        if (Inform05 == "Y")
                                        {
                                            if (PSH_Globals.SBO_Application.MessageBox("주 52시간 연장초과 했습니다. 그래도 입력하시겠습니까? ?", 2, "예", "아니오") == 1)
                                            {
                                                OverTime = 0;
                                            }
                                        }
                                        else
                                        {
                                            errNum = 1;
                                            throw new Exception();
                                        }
                                    }
                                    
                                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52) //스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                                    {
                                        if (LateTo + EarlyOff + GoOut > 0.1) //지각, 조퇴, 외출시간이 있으면 교대일수 없다.
                                        {
                                            Rotation = 0;
                                        }
                                        else
                                        {
                                            Rotation = 1;
                                        }

                                        if (codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11")
                                        {
                                            Rotation = 0;
                                        }

                                        if ((JIGTYP == "02" || JIGTYP == "03" || JIGCOD == "73") && PAYTYP == "2")
                                        {
                                            Extend = PH_PY008_Extend15(MSTCOD);

                                            SpecialTot = PH_PY008_Special(MSTCOD); //총특근시간

                                            if (SpecialTot >= 24)
                                            {
                                                Special = 0;
                                            }
                                            else if ((Special + SpecialTot) >= 24)
                                            {
                                                Special -= (Special + SpecialTot - 24);
                                            }
                                        }
                                        else
                                        {
                                            Extend = System.Convert.ToDouble(oForm.Items.Item("Extend").Specific.Value.ToString().Trim());
                                        }

                                        if (WorkType == "A01" || WorkType == "A02" || codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11") //무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                                        {
                                            DangerCD = "";
                                        }
                                        else
                                        {
                                            if ((JIGTYP == "04" || JIGTYP == "05") && DangerCD == "" && JIGCOD != "73") //전문직, 계약직 이며 위해코드가 없으면 위해코드를 기타로.. //JIGCOD 73 추가 20180201 황영수 사무계약직의 경우 추가 안되도록 수정 김택근과장 요청
                                            {
                                                if (CLTCOD == "1")
                                                {
                                                    DangerCD = "31";
                                                }
                                            }
                                        }

                                        if (Base + Special > 0) //근무시간이 있으나 4시간미만이면 위해수당 없다.
                                        {
                                            if (Special + SpExtend + Base + Extend < 4)
                                            {
                                                DangerCD = "";
                                            }
                                        }
                                        else
                                        {   
                                            if (CLTCOD == "1" && JIGCOD != "73") //근무시간이 없으면 위해코드를 기타로. //JIGCOD 73 추가 20180201 황영수 사무계약직의 경우 추가 안되도록 수정 김택근과장 요청
                                            {
                                                if (DangerCD.Trim() != "")
                                                {
                                                    DangerCD = "31";
                                                }
                                            }
                                        }
                                        
                                        if (DangerCD == "")
                                        {
                                            DangerNu = 0;
                                        }
                                        else
                                        {
                                            DangerNu = 1;
                                        }

                                        sQry = "Update ZPH_PY008";
                                        sQry += " Set ShiftDat = '" + ShiftDat + "',";
                                        sQry += " GNMUJO = '" + GNMUJO + "',";
                                        sQry += " WorkType = '" + WorkType + "',";
                                        sQry += " GetDate = '" + GetDate + "',";
                                        sQry += " GetTime = " + GetTime + ",";
                                        sQry += " OffDate = '" + OffDate + "',";
                                        sQry += " OffTime = " + OffTime + ",";
                                        sQry += " Base = " + Base + ",";
                                        sQry += " Extend = " + Extend + ",";
                                        sQry += " Special = " + Special + ",";
                                        sQry += " SpExtend = " + SpExtend + ",";
                                        sQry += " Midnight = " + Midnight + ",";
                                        sQry += " LateTo = " + LateTo + ",";
                                        sQry += " EarlyOff = " + EarlyOff + ",";
                                        sQry += " EarlyTo = " + EarlyTo + ",";
                                        sQry += " SEarlyTo = " + SEarlyTo + ",";
                                        sQry += " EducTran = " + EducTran + ",";
                                        sQry += " GoOut = " + GoOut + ",";
                                        sQry += " GOFrTime = " + GOFrTime + ",";
                                        sQry += " GOToTime = " + GOToTime + ",";
                                        sQry += " GOFrTim2 = " + GOFrTim2 + ",";
                                        sQry += " GOToTim2 = " + GoToTim2 + ",";
                                        sQry += " ActCode = '" + ActCode + "',";
                                        sQry += " ActText = '" + ActText + "',";
                                        sQry += " DangerCD = '" + DangerCD + "',";
                                        sQry += " DangerNu = " + DangerNu + ",";
                                        sQry += " Rotation = " + Rotation + ",";
                                        sQry += " RToDate = '" + RToDate + "',";
                                        sQry += " RToTime = " + RToTime + ",";
                                        sQry += " ROffDate = '" + ROffDate + "',";
                                        sQry += " ROffTime = " + ROffTime + ",";
                                        sQry += " ROutTime = " + ROutTime + ",";
                                        sQry += " RInTime = " + RInTime + ",";
                                        sQry += " Confirm = '" + Confirm + "',";
                                        sQry += " Comment = '" + Comment + "',";
                                        sQry += " RotateYN = '" + RotateYN + "',";
                                        sQry += " DAYTyp = '" + DAYTyp + "'";
                                        sQry += " Where CLTCOD = '" + CLTCOD + "'";
                                        sQry += " And PosDate = '" + PosDate + "'";
                                        sQry += " And MSTCOD = '" + MSTCOD + "'";

                                        oRecordSet.DoQuery(sQry);
                                    }
                                }
                            }

                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + oDS_PH_PY008.Rows.Count + "건 저장중...!";

                            PH_PY008_LeaveOfAbsence(CLTCOD, PosDate, MSTCOD); // 휴직자 확인 후 휴직일 경우 자료 입력
                        }
                    }
                }
                else //근태이상자 적용
                {
                    //사원의 직무와 위해 조회
                    sQry = "  select  Activity = a.U_Activity,";
                    sQry += "         ActText = b.CodeNm,";
                    sQry += "         DangerCD = b.DangerCD,";
                    sQry += "         JIGTYP = a.U_JIGTYP,";
                    sQry += "         PAYTYP = a.U_PAYTYP, ";
                    sQry += "         JIGCOD = a.U_JIGCOD ";
                    sQry += " FROM    [@PH_PY001A] AS a";
                    sQry += "         Left Join";
                    sQry += "         (";
                    sQry += "             Select  Code = t1.U_Code,";
                    sQry += "                     CodeNm = t1.U_CodeNm,";
                    sQry += "                     DangerCD = t1.U_Char1 ";
                    sQry += "             FROM    [@PS_HR200H] t";
                    sQry += "                     Inner Join";
                    sQry += "                     [@PS_HR200L] t1";
                    sQry += "                         On t.Code = t1.Code ";
                    sQry += "             Where   t1.Code = 'P127'";
                    sQry += "                     and U_Char2 = '" + CLTCOD + "'";
                    sQry += "         ) AS b";
                    sQry += "             On a.U_Activity = b.Code";
                    sQry += " Where   a.Code = '" + MSTCOD + "'";

                    oRecordSet.DoQuery(sQry);

                    ActCode = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    ActText = oRecordSet.Fields.Item(1).Value.ToString().Trim();
                    DangerCD = oRecordSet.Fields.Item(2).Value.ToString().Trim();
                    JIGTYP = oRecordSet.Fields.Item(3).Value.ToString().Trim(); //직원구분
                    PAYTYP = oRecordSet.Fields.Item(4).Value.ToString().Trim(); //급여구분
                    JIGCOD = oRecordSet.Fields.Item(5).Value.ToString().Trim(); //직급

                    if (oForm.Items.Item("SDay").Specific.Value == "월") //월요일은 제외
                    {
                        OverTime = 0;
                    }
                    else
                    {
                        OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD);
                    }

                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime >= 52)
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("주 52시간을 초과 했습니다. 그래도 입력하시겠습니까? ?", 2, "예", "아니오") == 1)
                        {
                            OverTime = 0;
                        }
                    }

                    if (Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52)
                    {
                        if (LateTo + EarlyOff + GoOut > 0.1) //지각, 조퇴, 외출시간이 있으면 교대일수 없다.
                        {
                            Rotation = 0;
                        }
                        else
                        {
                            Rotation = 1;
                        }
                        
                        if (codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11")
                        {
                            Rotation = 0;
                        }

                        if ((JIGTYP == "02" || JIGTYP == "03" || JIGCOD == "73") && PAYTYP == "2")
                        {
                            Extend = PH_PY008_Extend15(MSTCOD);
                        }
                        else
                        {
                            Extend = System.Convert.ToDouble(oForm.Items.Item("Extend").Specific.Value.ToString().Trim());
                        }
                        

                        if (CLTCOD == "1")
                        {
                            if (WorkType == "A01" || WorkType == "A02" || codeHelpClass.Left(WorkType, 1) == "F" || WorkType == "D11") //무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                            {
                                DangerCD = "";
                            }
                        }
                        else
                        {
                            if (codeHelpClass.Left(WorkType, 1) != "H" && codeHelpClass.Left(WorkType, 1) != "C" && Special + SpExtend + Base + Extend >= 4)
                            {
                            }
                            else
                            {
                                DangerCD = "";
                            }
                        }
                        
                        sQry = "Update ZPH_PY008";
                        sQry += " Set ShiftDat = '" + ShiftDat + "',";
                        sQry += " GNMUJO = '" + GNMUJO + "',";
                        sQry += " WorkType = '" + WorkType + "',";
                        sQry += " GetDate = '" + GetDate + "',";
                        sQry += " GetTime = " + GetTime + ",";
                        sQry += " OffDate = '" + OffDate + "',";
                        sQry += " OffTime = " + OffTime + ",";
                        sQry += " Base = " + Base + ",";
                        sQry += " Extend = " + Extend + ",";
                        sQry += " Special = " + Special + ",";
                        sQry += " SpExtend = " + SpExtend + ",";
                        sQry += " Midnight = " + Midnight + ",";
                        sQry += " LateTo = " + LateTo + ",";
                        sQry += " EarlyOff = " + EarlyOff + ",";
                        sQry += " EarlyTo = " + EarlyTo + ",";
                        sQry += " SEarlyTo = " + SEarlyTo + ",";
                        sQry += " EducTran = " + EducTran + ",";
                        sQry += " GoOut = " + GoOut + ",";
                        sQry += " GOFrTime = " + GOFrTime + ",";
                        sQry += " GOToTime = " + GOToTime + ",";
                        sQry += " GOFrTim2 = " + GOFrTim2 + ",";
                        sQry += " GOToTim2 = " + GoToTim2 + ",";
                        sQry += " Rotation = " + Rotation + ",";
                        sQry += " RToDate = '" + RToDate + "',";
                        sQry += " RToTime = " + RToTime + ",";
                        sQry += " ROffDate = '" + ROffDate + "',";
                        sQry += " ROffTime = " + ROffTime + ",";
                        sQry += " ROutTime = " + ROutTime + ",";
                        sQry += " RInTime = " + RInTime + ",";
                        sQry += " Confirm = '" + Confirm + "'";
                        sQry += " Where CLTCOD = '" + CLTCOD + "'";
                        sQry += " And PosDate = '" + PosDate + "'";
                        sQry += " And MSTCOD = '" + MSTCOD + "'";

                        oRecordSet.DoQuery(sQry);
                    }
                }

                oForm.Items.Item("WorkType").Specific.Select(oWorkType, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch(Exception ex)
            {
                ProgressBar01.Stop();

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("주 52시간 연장초과 했습니다. 입력할 수 없습니다. 현재 총 연장시간은 " + Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime + "시간 입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 11)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근무조를 입력하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgressBar01.Stop();

                if (oForm.DataSources.UserDataSources.Item("Chk").Value.ToString().Trim() == "N")
                {
                    PH_PY008_MTX01("A");
                }
                else
                {
                    PH_PY008_MTX01("S");
                }

                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 휴직자 체크후 변경
        /// </summary>
        /// <param name="CLTCODE">사업장</param>
        /// <param name="DocDate">날짜</param>
        /// <param name="MSTCOD">사번</param>
        /// <returns></returns>
        private void PH_PY008_LeaveOfAbsence(string CLTCODE, string DocDate, String MSTCOD )
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sQry = "Exec [PH_PY008_03] '" + CLTCODE + "','" + DocDate + "','" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 주근무시간 체크
        /// </summary>
        /// <param name="PosDate">일자</param>
        /// <param name="MSTCOD">사원성명</param>
        /// <returns></returns>
        private double PH_PY008_OverTimeCheck(string PosDate, string MSTCOD)
        {
            string sQry;
            string FrDate;
            string ToDate;
            string CLTCOD;
            string OverTimeYN;
            short errNum = 0;
            double returnValue = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;

                if (oForm.Items.Item("SDay").Specific.Value == "일")
                {
                    ToDate = PosDate;

                    //월요일 찾기
                    sQry = "  SELECT  U_Date";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 2 ";
                    sQry += "         and U_Date between DATEADD(dd, -6, '" + PosDate + "') and '" + PosDate + "'";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        FrDate = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");
                    }
                    else
                    {
                        errNum = 1;
                        throw new Exception();
                    }
                }
                else if (oForm.Items.Item("SDay").Specific.Value == "월")
                {
                    FrDate = PosDate;

                    //일요일 찾기
                    sQry = "  SELECT  U_Date";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 1 ";
                    sQry += "         and U_Date between '" + PosDate + "' and DATEADD(dd, 6, '" + PosDate + "')";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        ToDate = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");
                    }
                    else
                    {
                        errNum = 2;
                        throw new Exception();
                    }
                }
                else
                {
                    //월요일 찾기
                    sQry = "  SELECT  U_Date";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 2 ";
                    sQry += "         and U_Date between DATEADD(dd, -6, '" + PosDate + "') and '" + PosDate + "'";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        FrDate = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");
                    }
                    else
                    {
                        errNum = 1;
                        throw new Exception();
                    }

                    //일요일 찾기
                    sQry = "  SELECT  U_Date ";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 1 ";
                    sQry += "         and U_Date between '" + PosDate + "' and  DATEADD(dd, 6, '" + PosDate + "')";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        ToDate = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");
                    }
                    else
                    {
                        errNum = 2;
                        throw new Exception();
                    }
                }

                sQry = "SELECT Isnull(U_Inform05,'N') From [@PH_PY001A] Where Code = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                OverTimeYN = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                // 기존 로직에서 Base Special 추가로 입력해서 주 52시간 초과인지 계산함.
                sQry = "  Select  Sum(Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo)";
                sQry += " From    ZPH_PY008";
                sQry += " Where   PosDate Between '" + FrDate + "' and '" + ToDate + "'";
                sQry += "         And MSTCOD = '" + MSTCOD + "'";
                sQry += "         And PosDate <> '" + PosDate + "'";

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 3;
                    throw new Exception();
                }
                else if (OverTimeYN == "Y")
                {
                    returnValue = -100;
                }
                else
                {
                    returnValue = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
            }
            catch(Exception ex)
            {
                returnValue = 15;

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("[12시간초과체크] 기준일자(월요일)을 가져오지 못했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("[12시간초과체크] 기준일자(일요일)을 가져오지 못했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum ==3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 주 근무시간 월급제 사원 2.5 계산을 위한 신규오버타임
        /// </summary>
        /// <param name="PosDate"></param>
        /// <param name="MSTCOD"></param>
        /// <returns></returns>
        private double PH_PY008_OverTimeCheck2(string PosDate, string MSTCOD)
        {
            string sQry;
            string FrDate;
            string ToDate;
            string CLTCOD;
            string OverTimeYN;
            short errNum = 0;
            double returnValue = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;

                if (oForm.Items.Item("SDay").Specific.Value == "일")
                {
                    ToDate = PosDate;

                    //월요일 찾기
                    sQry = "  SELECT  U_Date";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 2";
                    sQry += "         and U_Date between DATEADD(dd, -6, '" + PosDate + "') and '" + PosDate + "'";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        FrDate = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    }
                    else
                    {
                        errNum = 1;
                        throw new Exception();
                    }
                }
                else if (oForm.Items.Item("SDay").Specific.Value == "월")
                {
                    FrDate = PosDate;

                    //일요일 찾기
                    sQry = "  SELECT  U_Date ";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 1 ";
                    sQry += "         and U_Date between '" + PosDate + "' and DATEADD(dd, 6, '" + PosDate + "')";

                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount > 0)
                    {
                        ToDate = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    }
                    else
                    {
                        errNum = 2;
                        throw new Exception();
                    }
                }
                else
                {
                    //월요일 찾기
                    sQry = "  SELECT  U_Date ";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 2 ";
                    sQry += "         and U_Date between DATEADD(dd, -6, '" + PosDate + "') and '" + PosDate + "'";

                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount > 0)
                    {
                        FrDate = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    }
                    else
                    {
                        errNum = 1;
                        throw new Exception();
                    }

                    //일요일 찾기
                    sQry = "  SELECT  U_Date ";
                    sQry += " From    [@PH_PY003A] a";
                    sQry += "         Inner join";
                    sQry += "         [@PH_PY003B] b";
                    sQry += "             On a.Code = b.Code";
                    sQry += "             And a.U_CLTCOD = '" + CLTCOD + "'";
                    sQry += " Where   DatePart(Weekday, U_Date) = 1 ";
                    sQry += "         and U_Date between '" + PosDate + "' and  DATEADD(dd, 6, '" + PosDate + "')";

                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount > 0)
                    {
                        ToDate = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                    }
                    else
                    {
                        errNum = 2;
                        throw new Exception();
                    }
                }

                sQry = "SELECT Isnull(U_Inform05,'N') From [@PH_PY001A] Where Code = '" + MSTCOD + "'";
                oRecordSet.DoQuery(sQry);

                OverTimeYN = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                //기존 로직에서 Base Special 추가로 입력해서 주 52시간 초과인지 계산함.
                sQry = "  Select  Sum(Extend + SpExtend + EarlyTo + SEarlyTo)";
                sQry += " From    ZPH_PY008";
                sQry += " Where   PosDate Between '" + FrDate + "' and '" + ToDate + "'";
                sQry += "         And MSTCOD = '" + MSTCOD + "'";
                sQry += "         And PosDate <> '" + PosDate + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 3;
                }
                else if (OverTimeYN == "Y")
                {
                    returnValue = -100;
                }
                else
                {
                    returnValue = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
            }
            catch(Exception ex)
            {
                returnValue = 15;

                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("[12시간초과체크] 기준일자(월요일)을 가져오지 못했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("[12시간초과체크] 기준일자(일요일)을 가져오지 못했습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }
        
        /// <summary>
        /// 사무기술직 월 15시간 => 20시간으로 변경 2014/07/01부 N.G.Y
        /// </summary>
        /// <param name="MSTCOD"></param>
        /// <returns></returns>
        private double PH_PY008_Extend15(string MSTCOD)
        {
            string sQry;
            string CLTCOD;
            string YM;
            double StTime;
            double OverTime;
            double CLTCODTIME;
            double returnValue = 0;
            short errNum = 0;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                OverTime = PH_PY008_OverTimeCheck2(oForm.Items.Item("PosDate").Specific.Value, MSTCOD);

                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;
                YM = codeHelpClass.Left(oForm.Items.Item("PosDate").Specific.Value, 6);

                if (CLTCOD == "1")
                {
                    CLTCODTIME = 2.5;
                }
                else
                {
                    CLTCODTIME = 2.0;
                }

                if (OverTime + CLTCODTIME > 12)
                {
                    CLTCODTIME = 0.0;
                }

                if (oForm.Items.Item("SDayOff").Specific.Value == "1")
                {
                    StTime = 20; //월 20시간

                    sQry = "  Select  Isnull(Sum(Extend),0)";
                    sQry += " From    ZPH_PY008";
                    sQry += " Where   Convert(char(6),PosDate,112) = '" + YM + "'";
                    sQry += "         And CLTCOD = '" + CLTCOD + "'";
                    sQry += "         AND MSTCOD = '" + MSTCOD + "'";

                    oRecordSet.DoQuery(sQry);

                    if (oRecordSet.RecordCount > 0)
                    {
                        if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) - StTime < 0 && Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) - StTime >= CLTCODTIME)
                        {
                            returnValue = CLTCODTIME;
                        }
                        else if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) - StTime >= 0)
                        {
                            returnValue = 0;
                        }
                        else
                        {
                            returnValue = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) - StTime;
                        }
                    }
                    else
                    {
                        errNum = 1;
                        throw new Exception();
                    }
                }
            }
            catch(Exception ex)
            {
                if(errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사무기술직(월급제)의 연장근무시간 계산이 되지않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 사무기술직 특근시간 월 24시간 초과 할 수 없음. 2014/07/01부 N.G.Y
        /// </summary>
        /// <param name="MSTCOD"></param>
        /// <returns></returns>
        private double PH_PY008_Special(string MSTCOD)
        {
            string sQry;
            string CLTCOD;
            string YM;
            double returnValue = 0;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                YM = codeHelpClass.Left(oForm.Items.Item("PosDate").Specific.Value.ToString().Trim(), 6);

                sQry = "  Select  Isnull(Sum(Special),0)";
                sQry += " From    ZPH_PY008";
                sQry += " Where   Convert(char(6),PosDate,112) = '" + YM + "'";
                sQry += "         And CLTCOD = '" + CLTCOD + "'";
                sQry += "         AND MSTCOD = '" + MSTCOD + "'";

                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount > 0)
                {
                    returnValue = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
                else
                {
                    returnValue = 0;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 근무시간 계산
        /// </summary>
        /// <param name="OffTime"></param>
        private void PH_PY008_Time_Calc_Main(string OffTime)
        {
            short i;
            string CLTCOD;
            //20190320 황영수 수정 S
            string DayType;
            //20190320 황영수 수정 E
            string ShiftDat; //근무형태
            string GNMUJO; //근무조
            string DayOff; //평일휴일구분
            string PosDate; //기준일
            string GetDate; //출근일
            string OffDate; //퇴근일
            string GetTime; //출근시간
            string FromTime;
            string ToTime;
            double hTime1; //오전10분휴식시간
            double hTime5; //야간휴식시간
            string NextDay;
            string TimeType;
            string sQry;
            string STime = string.Empty;
            string ETime = string.Empty;
            double EarlyTo = 0;
            double SEarlyTo = 0;
            double Base = 0;
            double Special = 0;
            double Extend = 0;
            double SpExtend = 0;
            double Midnight = 0;
            string WorkType;
            string Team1;
            string Team2;
            string Team3;
            short errNum = 0;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                ShiftDat = oForm.Items.Item("ShiftDat").Specific.Value.ToString().Trim();
                GNMUJO = oForm.Items.Item("GNMUJO").Specific.Value.ToString().Trim();

                PosDate = oForm.Items.Item("PosDate").Specific.Value.ToString().Trim();
                GetDate = oForm.Items.Item("GetDate").Specific.Value.ToString().Trim();
                OffDate = oForm.Items.Item("OffDate").Specific.Value.ToString().Trim();

                //20190320 황영수 수정 S
                sQry = "  select  case";
                sQry += "             when b.U_WorkType2 ='A00' then '1'";
                sQry += "             else '2'";
                sQry += "         end as Worktype";
                sQry += " from    [@PH_PY003A] a";
                sQry += "         inner join";
                sQry += "         [@PH_PY003B] b";
                sQry += "             on a.Code = b.Code ";
                sQry += " Where   a.U_CLTCOD = '" + CLTCOD + "'";
                sQry += "         and b.U_Date = '" + PosDate + "'";

                oRecordSet.DoQuery(sQry);

                DayType = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                if (oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() == "4" && DayType == "1")
                {
                    DayOff = "1";
                }
                else
                {
                    DayOff = oForm.Items.Item("DayOff1").Specific.Value.ToString().Trim();
                }
                //20190320 황영수 수정 E

                GetTime = oForm.Items.Item("GetTime").Specific.Value.ToString().Trim();
                OffTime = oForm.Items.Item("OffTime").Specific.Value.ToString().Trim();

                Team1 = oForm.Items.Item("Team1").Specific.Value.ToString().Trim();
                Team2 = oForm.Items.Item("Team2").Specific.Value.ToString().Trim();
                Team3 = oForm.Items.Item("Team3").Specific.Value.ToString().Trim();

                GetTime = "0" + GetTime;
                GetTime = codeHelpClass.Right(GetTime, 4);

                OffTime = "0" + OffTime;
                OffTime = codeHelpClass.Right(OffTime, 4);

                WorkType = oForm.Items.Item("WorkType").Specific.Value.ToString().Trim();

                if (GNMUJO == "41")
                {
                    if (Team1 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team1 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                if (GNMUJO == "42")
                {
                    if (Team2 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team2 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                if (GNMUJO == "43")
                {
                    if (Team3 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team3 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                //위 근무조를 받아와서 21 OR 22 로 변경함 (주간, 야간 시간표때문 PHPY002) (주52)
                if (ShiftDat == "4")
                {
                    // 20190320 황영수 수정 S
                    ShiftDat = "2";

                    sQry = "  Select  U_TimeType,";
                    sQry += "         U_FromTime,";
                    sQry += "         U_ToTime,";
                    sQry += "         U_NextDay";
                    sQry += " from    [@PH_PY002A] a";
                    sQry += "         Inner Join";
                    sQry += "         [@PH_PY002B] b";
                    sQry += "             On a.Code = b.Code ";
                    sQry += " Where   a.U_CLTCOD = '" + CLTCOD + "'"; //사업부
                    sQry += "         and a.U_SType = '" + ShiftDat + "'"; //교대
                    sQry += "         and a.U_Shift = '" + GNMUJO + "'"; //조
                    sQry += "         and b.U_DayType = '" + DayOff + "'"; //평일

                    ShiftDat = "4";
                }
                else
                {
                    sQry = "  Select  U_TimeType,";
                    sQry += "         U_FromTime,";
                    sQry += "         U_ToTime,";
                    sQry += "         U_NextDay";
                    sQry += " from    [@PH_PY002A] a";
                    sQry += "         Inner Join";
                    sQry += "         [@PH_PY002B] b";
                    sQry += "             On a.Code = b.Code ";
                    sQry += " Where   a.U_CLTCOD = '" + CLTCOD + "'"; //사업부
                    sQry += "         and a.U_SType = '" + ShiftDat + "'"; //교대
                    sQry += "         and a.U_Shift = '" + GNMUJO + "'"; //조
                    sQry += "         and b.U_DayType = '" + DayOff + "'"; //평일
                }
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    FromTime = oRecordSet.Fields.Item(1).Value.ToString().Trim();
                    FromTime = "0000" + FromTime;
                    FromTime = codeHelpClass.Right(FromTime, 4);

                    ToTime = oRecordSet.Fields.Item(2).Value.ToString().Trim();
                    ToTime = "0000" + ToTime;
                    ToTime = codeHelpClass.Right(ToTime, 4);

                    NextDay = oRecordSet.Fields.Item(3).Value.ToString().Trim();
                    TimeType = oRecordSet.Fields.Item(0).Value.ToString().Trim();

                    if (NextDay == "")
                    {
                        NextDay = "N";
                    }

                    if (NextDay == "N")
                    {
                        if (ToTime == "0000")
                        {
                            ToTime = "2400";
                        }
                    }

                    switch (TimeType)
                    {
                        case "40": //조출

                            if (GetDate == PosDate)
                            {
                                if (Convert.ToInt16(GetTime) < Convert.ToInt16(ToTime))
                                {
                                    STime = GetTime;
                                    ETime = ToTime;

                                    if (DayOff == "1")
                                    {
                                        EarlyTo = PH_PY008_Time_Calc(STime, ETime);
                                    }
                                    else
                                    {
                                        SEarlyTo = PH_PY008_Time_Calc(STime, ETime);
                                    }
                                }
                            }

                            break;

                        case "10":
                        case "50": //정상근무시간

                            if (GNMUJO == "11" || GNMUJO == "21" || GNMUJO == "31" || GNMUJO == "32")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일
                                        if (PosDate == GetDate)
                                        {
                                            //1교대1조, 2교대 1조당일
                                            if (Convert.ToInt16(GetTime) < Convert.ToInt16(FromTime))
                                            {
                                                STime = FromTime; //시작시간
                                            }
                                            else
                                            {
                                                STime = GetTime; //출근시간
                                            }
                                            
                                            if (GetDate != OffDate)
                                            {
                                                ETime = ToTime; //종료시간
                                            }
                                            else if (Convert.ToInt16(OffTime) < Convert.ToInt16(ToTime))
                                            {
                                                ETime = OffTime; //퇴근시간
                                            }
                                            else
                                            {
                                                ETime = ToTime; //종료시간
                                            }

                                            if (Convert.ToInt16(DayOff) == 1)
                                            {
                                                Base += PH_PY008_Time_Calc(STime, ETime); //평일
                                            }
                                            else
                                            {
                                                Special += PH_PY008_Time_Calc(STime, ETime); //휴일
                                            }
                                        }

                                        break;

                                    case "Y": //익일
                                        break;
                                }
                            }
                            else if (GNMUJO == "22" || GNMUJO == "33")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일
                                        if (PosDate == GetDate)
                                        {
                                            if (Convert.ToInt16(GetTime) < Convert.ToInt16(FromTime))
                                                STime = FromTime; //시작시간
                                            else
                                                STime = GetTime; //출근시간

                                            if (GetDate == OffDate)
                                            {
                                                if (Convert.ToInt16(OffTime) < 2400)
                                                {
                                                    ETime = OffTime; //퇴근시간
                                                }
                                                else
                                                {
                                                    ETime = "2400";
                                                }
                                            }
                                            else
                                            {
                                                ETime = "2400";
                                            }

                                            if (Convert.ToInt16(DayOff) == 1)
                                            {
                                                Base += PH_PY008_Time_Calc(STime, ETime);
                                            }
                                            else
                                            {
                                                Special += PH_PY008_Time_Calc(STime, ETime);
                                            }
                                        }

                                        break;

                                    case "Y": //익일
                                        if (PosDate == OffDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            STime = "0000"; //시작시간은 00시

                                            if (Convert.ToInt16(OffTime) < Convert.ToInt16(ToTime))
                                            {
                                                ETime = OffTime; //퇴근시간
                                            }
                                            else
                                            {
                                                ETime = ToTime; //종료시간
                                            }
                                        }

                                        if (Convert.ToDouble(DayOff) == 1)
                                        {
                                            Base += PH_PY008_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            Special += PH_PY008_Time_Calc(STime, ETime);
                                        }
                                            
                                        break;
                                }
                            }

                            break;

                        case "65":
                        case "66":
                        case "15": //오전, 오후 휴식시간, 점심시간

                            if (GNMUJO == "11" || GNMUJO == "21" || GNMUJO == "22" || GNMUJO == "31" || GNMUJO == "32" || GNMUJO == "33")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일

                                        if (PosDate != GetDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else if (PosDate == OffDate)
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                if (Convert.ToDouble(FromTime) <= Convert.ToDouble(OffTime))
                                                {
                                                    STime = FromTime; //시작시간
                                                    if (Convert.ToDouble(ToTime) < Convert.ToDouble(OffTime))
                                                    {
                                                        ETime = ToTime;
                                                    }
                                                    else
                                                    {
                                                        ETime = OffTime;
                                                    }
                                                }
                                                else
                                                {
                                                    STime = "0000";
                                                    ETime = "0000";
                                                }
                                            }
                                            else if (Convert.ToDouble(GetTime) > Convert.ToDouble(ToTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = GetTime; //출근시간
                                                if (Convert.ToDouble(ToTime) < Convert.ToDouble(OffTime))
                                                {
                                                    ETime = ToTime; //종료시간
                                                }
                                                else
                                                {
                                                    ETime = OffTime;//퇴근시간
                                                }
                                            }
                                        }
                                        else if (Convert.ToDouble(ToTime) < Convert.ToDouble(GetTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime;
                                                ETime = ToTime;
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }
                                            ETime = ToTime;
                                        }

                                        break;

                                    case "Y": //익일

                                        if (PosDate == OffDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            STime = FromTime;
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }
                                        }

                                        break;
                                }
                            }

                            hTime1 = PH_PY008_Time_Calc(STime, ETime); //오전휴식시간

                            if (GNMUJO == "22" && TimeType == "65")
                            {
                                Midnight -= hTime1;
                            }
                            else if (DayOff == "1")
                            {
                                Base -= hTime1;
                            }
                            else
                            {
                                Special -= hTime1;
                            }
                                
                            break;
                            
                        case "20":
                        case "60": //연장근무

                            if (GNMUJO == "11" || GNMUJO == "21" || GNMUJO == "31" || GNMUJO == "32")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일

                                        if (PosDate != OffDate)
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime;
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }
                                                
                                            ETime = "2400";
                                        }
                                        else if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            STime = FromTime; //종료시간
                                            ETime = OffTime; //퇴근시간
                                        }

                                        if (DayOff == "1")
                                        {
                                            Extend += PH_PY008_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend += PH_PY008_Time_Calc(STime, ETime);
                                        }

                                        break;

                                    case "Y": //익일

                                        if (PosDate != OffDate)
                                        {
                                            STime = "0000";
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }
                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }

                                        if (DayOff == "1")
                                        {
                                            Extend += PH_PY008_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend += PH_PY008_Time_Calc(STime, ETime);
                                        }

                                        break;
                                }
                            }
                            else if (GNMUJO == "22" || GNMUJO == "33")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일
                                        break;

                                    case "Y": //익일

                                        if (PosDate != OffDate)
                                        {
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                STime = FromTime;
                                                ETime = OffTime;
                                            }
                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }

                                        if (DayOff == "1")
                                        {
                                            Extend += PH_PY008_Time_Calc(STime, ETime);
                                        }
                                        else
                                        {
                                            SpExtend += PH_PY008_Time_Calc(STime, ETime);
                                        }

                                        break;
                                }
                            }

                            break;

                        case "25": //저녁휴식
                            break;

                        case "30": //심야시간
                        
                            if (GNMUJO == "11" || GNMUJO == "21" || GNMUJO == "22" || GNMUJO == "31" || GNMUJO == "32" || GNMUJO == "33")
                            {
                                switch (NextDay)
                                {
                                    case "N": //당일

                                        if (PosDate != OffDate)
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime; //시작시간
                                            }
                                            else
                                            {
                                                STime = GetTime; //출근시간
                                            }

                                            ETime = "2400";
                                        }
                                        else if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime; //시작시간
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }

                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime; //종료시간
                                            }
                                        }

                                        Midnight += PH_PY008_Time_Calc(STime, ETime);

                                        break;

                                    case "Y": //익일

                                        if (PosDate != OffDate)
                                        {
                                            STime = "0000";
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime; //퇴근시간
                                            }
                                            else
                                            {
                                                ETime = ToTime; //종료시간
                                            }
                                        }
                                        else
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        Midnight += PH_PY008_Time_Calc(STime, ETime);
                                        break;
                                }
                            }

                            break;

                        case "35": //야간휴식 수정해야함
                            {
                                //hTime5
                                switch (NextDay)
                                {
                                    case "N": //당일

                                        if (PosDate == OffDate)
                                        {
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = "0000";
                                                ETime = "0000";
                                            }
                                            else
                                            {
                                                if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                                {
                                                    STime = FromTime;
                                                }
                                                else
                                                {
                                                    STime = GetTime;
                                                }

                                                if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                                {
                                                    ETime = OffTime;
                                                }
                                                else
                                                {
                                                    ETime = ToTime;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //다음날 퇴근
                                            if (Convert.ToDouble(GetTime) < Convert.ToDouble(FromTime))
                                            {
                                                STime = FromTime;
                                            }
                                            else
                                            {
                                                STime = GetTime;
                                            }

                                            ETime = ToTime;
                                        }

                                        break;

                                    case "Y":

                                        if (PosDate == OffDate)
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else if (Convert.ToDouble(OffTime) < Convert.ToDouble(FromTime))
                                        {
                                            STime = "0000";
                                            ETime = "0000";
                                        }
                                        else
                                        {
                                            STime = FromTime;
                                            if (Convert.ToDouble(OffTime) < Convert.ToDouble(ToTime))
                                            {
                                                ETime = OffTime;
                                            }
                                            else
                                            {
                                                ETime = ToTime;
                                            }
                                        }

                                        break;
                                }

                                hTime5 = PH_PY008_Time_Calc(STime, ETime);

                                if (DayOff == "1") //평일
                                {
                                    if (GNMUJO == "22")
                                    {
                                        Base -= hTime5; //기본근무
                                        Midnight -= hTime5; //심야시간에서 차감
                                    }
                                    else if(GNMUJO == "33")    // 3교대 3조는 연장근무에서 차감하면 안되기 때문에 별도 로직 적용(2020.01.30 황영수 S)
                                    {
                                        Midnight -= hTime5; 
                                    }                         // 3교대 3조는 연장근무에서 차감하면 안되기 때문에 별도 로직 적용(2020.01.30 황영수 E)
                                    else
                                    {
                                        Extend -= hTime5; //연장근무에서 차감
                                        Midnight -= hTime5; //심야시간에서 차감
                                    }
                                }
                                else //휴일
                                {
                                    if (GNMUJO == "22")
                                    {
                                        Special -= hTime5;
                                        Midnight -= hTime5; //심야시간에서 차감
                                    }
                                    else if (GNMUJO == "33")          // 3교대 3조는 연장근무에서 차감하면 안되기 때문에 별도 로직 적용(2020.01.30 황영수 S)
                                    {
                                        Midnight -= hTime5; 
                                    }                                 // 3교대 3조는 연장근무에서 차감하면 안되기 때문에 별도 로직 적용(2020.01.30 황영수 E)
                                    else
                                    {
                                        //다음날 퇴근은 연장근무임
                                        SpExtend -= hTime5; //연장근무에서 차감
                                        Midnight -= hTime5; //심야시간에서 차감
                                    }
                                }
                                
                                break;
                            }
                    }
                    oRecordSet.MoveNext();
                }

                oForm.Items.Item("EarlyTo").Specific.Value = PH_PY008_hhmm_Calc(EarlyTo, ""); //조출
                oForm.Items.Item("Base").Specific.Value = PH_PY008_hhmm_Calc(Base, WorkType); //기본
                oForm.Items.Item("Extend").Specific.Value = PH_PY008_hhmm_Calc(Extend, ""); //연장
                oForm.Items.Item("Midnight").Specific.Value = PH_PY008_hhmm_Calc(Midnight, ""); //심야
                oForm.Items.Item("SEarlyTo").Specific.Value = PH_PY008_hhmm_Calc(SEarlyTo, ""); //특근조출
                oForm.Items.Item("Special").Specific.Value = PH_PY008_hhmm_Calc(Special, ""); //특근
                oForm.Items.Item("SpExtend").Specific.Value = PH_PY008_hhmm_Calc(SpExtend, ""); //특근연장
            }
            catch(Exception ex)
            {
                if(errNum == 1)
                {
                    //처리 없이 메소드 종료
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 시작시각-종료시각
        /// </summary>
        /// <param name="GetTime"></param>
        /// <param name="OffTime"></param>
        /// <returns></returns>
        private double PH_PY008_Time_Calc(string GetTime, string OffTime)
        {
            int STime;
            int ETime;
            double returnValue = 0;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                STime = Convert.ToInt32(codeHelpClass.Mid(GetTime, 0, 2)) * 3600 + Convert.ToInt32(codeHelpClass.Mid(GetTime, 2, 2)) * 60;
                ETime = Convert.ToInt32(codeHelpClass.Mid(OffTime, 0, 2)) * 3600 + Convert.ToInt32(codeHelpClass.Mid(OffTime, 2, 2)) * 60;

                returnValue = ETime - STime;

                if (string.IsNullOrEmpty(returnValue.ToString().Trim()))
                {
                    returnValue = 0;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 근무조건의 시간 계산
        /// </summary>
        /// <param name="mTime"></param>
        /// <param name="pWorkType"></param>
        /// <returns></returns>
        private double PH_PY008_hhmm_Calc(double mTime, string pWorkType)
        {
            int hh;
            double MM;
            double returnValue = 0;

            try
            {
                hh = Convert.ToInt32(Math.Truncate(mTime / 3600));

                MM = (mTime % 3600) / 60;

                if (MM > 0)
                {
                    if (MM > 30)
                    {
                        MM = 1;
                    }
                    else
                    {
                        MM = 0.5;
                    }
                }
                else if (MM == 0)
                {
                    MM = 0;
                }
                else if (MM < -30)
                {
                    MM = -1;
                }
                else
                {
                    MM = -0.5;
                }

                returnValue = hh + MM;

                if (pWorkType == "D09" || pWorkType == "D10") // 근태구분이 반차일 경우 무조건 4시간 반환(2014.04.21 송명규 추가)
                {
                    returnValue = 4;
                }

                if (string.IsNullOrEmpty(returnValue.ToString().Trim()))
                {
                    returnValue = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 날자변경에 대한 자료 Set
        /// </summary>
        private void PH_PY008_ChDate_Set()
        {
            string sQry;
            string pDate;
            string CLTCOD;
            string sPosDate;
            short errNum = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;
                sPosDate = oForm.Items.Item("SPosDate").Specific.Value;

                //해당일의 기본근태
                //해당일의 근무일 등록 확인
                sQry = "  Select  U_WorkType";
                sQry += " From    [@PH_PY003A] AS a";
                sQry += "         Inner Join";
                sQry += "         [@PH_PY003B] AS b";
                sQry += "             On a.Code = b.Code ";
                sQry += " Where   a.U_CLTCOD = '" + CLTCOD + "'";
                sQry += "         And b.U_Date = '" + sPosDate + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0)
                {
                    errNum = 1;
                    throw new Exception();
                }
                else
                {
                    oWorkType = oRecordSet.Fields.Item(0).Value.ToString();
                    oForm.Items.Item("WorkType").Specific.Select(oRecordSet.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    PH_PY008_ChGNMUJO_Set();
                }

                oForm.DataSources.UserDataSources.Item("PosDate").Value = sPosDate;
                oForm.DataSources.UserDataSources.Item("GetDate").Value = sPosDate;

                pDate = PH_PY008_DaySelect(sPosDate);

                oForm.DataSources.UserDataSources.Item("SDay").Value = pDate;
                oForm.DataSources.UserDataSources.Item("Day1").Value = pDate;

                oForm.Items.Item("SDayOff").Specific.Select(PH_PY008_DayOffSelect(sPosDate), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("DayOff1").Specific.Select(PH_PY008_DayOffSelect(sPosDate), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근태월력의 근무일 등록을 하지않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 근무조 변경에 대한 자료 Set
        /// </summary>
        private void PH_PY008_ChGNMUJO_Set()
        {
            string sQry;
            string Team1;
            string Team2;
            string Team3;
            string GNMUJO;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                // 근무조를 받아와서 그 근무조가 야간인지 주간인지 휴무인지 체크함. (주52)
                // 근무조를 받아와서 체크해서 21 or 22로 변경하는 건 시간표(PH_PY003) 때문임.
                Team1 = oForm.Items.Item("Team1").Specific.Value.ToString().Trim();
                Team2 = oForm.Items.Item("Team2").Specific.Value.ToString().Trim();
                Team3 = oForm.Items.Item("Team3").Specific.Value.ToString().Trim();
                GNMUJO = oForm.Items.Item("GNMUJO").Specific.Value.ToString().Trim();

                if (GNMUJO == "41")
                {
                    if (Team1 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team1 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                if (GNMUJO == "42")
                {
                    if (Team2 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team2 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                if (GNMUJO == "43")
                {
                    if (Team3 == "D")
                    {
                        GNMUJO = "21";
                    }
                    else if (Team3 == "N")
                    {
                        GNMUJO = "22";
                    }
                    else
                    {
                        GNMUJO = "";
                    }
                }

                switch (GNMUJO)
                {
                    case "22":
                    case "33":
                    
                        sQry = "Select DateAdd(dd, 1, '" + oForm.Items.Item("SPosDate").Specific.Value + "')";
                        oRecordSet.DoQuery(sQry);
                        if (oRecordSet.RecordCount > 0)
                        {
                            oForm.DataSources.UserDataSources.Item("OffDate").Value = oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd");
                            oForm.DataSources.UserDataSources.Item("Day2").Value = PH_PY008_DaySelect(oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd"));
                            oForm.Items.Item("DayOff2").Specific.Select(PH_PY008_DayOffSelect(oRecordSet.Fields.Item(0).Value.ToString("yyyyMMdd")), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }

                        break;
                    
                    default:
                        
                        oForm.DataSources.UserDataSources.Item("OffDate").Value = oForm.Items.Item("SPosDate").Specific.Value;
                        oForm.DataSources.UserDataSources.Item("Day2").Value = PH_PY008_DaySelect(oForm.Items.Item("SPosDate").Specific.Value);
                        oForm.Items.Item("DayOff2").Specific.Select(PH_PY008_DayOffSelect(oForm.Items.Item("SPosDate").Specific.Value), SAPbouiCOM.BoSearchKey.psk_ByValue);

                        break;
                }

                if (oForm.Items.Item("GNMUJO").Specific.Value == "11" && oForm.Items.Item("SDayOff").Specific.Value == "1")
                {
                    //근무조가 1교대1조 평일일 경우
                    switch (oForm.Items.Item("SCLTCOD").Specific.Value)
                    {
                        case "1":
                        case "3":
                            
                            oForm.Items.Item("GetTime").Specific.Value = "0830";
                            oForm.Items.Item("OffTime").Specific.Value = "1730";

                            break;
                            
                        case "2":
                            
                            oForm.Items.Item("GetTime").Specific.Value = "0750";
                            oForm.Items.Item("OffTime").Specific.Value = "1700";
                            
                            break;
                    }
                }
                else
                {
                    oForm.Items.Item("GetTime").Specific.Value = "0000";
                    oForm.Items.Item("OffTime").Specific.Value = "0000";
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 선택된 사원 근태 자료 삭제
        /// </summary>
        private void PH_PY008_Delete()
        {
            string PosDate;
            string CLTCOD;
            string MSTCOD;
            string FullName;
            short i;
            string sQry;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
                FullName = oForm.Items.Item("FullName").Specific.Value.ToString().Trim();
                PosDate = oForm.Items.Item("PosDate").Specific.Value.ToString().Trim();

                if (PSH_Globals.SBO_Application.MessageBox("선택한사원 전체를 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
                {
                    if (oDS_PH_PY008.Rows.Count > 0)
                    {
                        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("삭제시작!", oDS_PH_PY008.Rows.Count, false);

                        oForm.Freeze(true);

                        for (i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                        {
                            if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value == "Y")
                            {
                                PosDate = Convert.ToDateTime(oDS_PH_PY008.Columns.Item("GetDate").Cells.Item(i).Value.ToString()).ToString("yyyyMMdd");
                                MSTCOD = oDS_PH_PY008.Columns.Item("MSTCOD").Cells.Item(i).Value.ToString().Trim();

                                sQry = "Delete From ZPH_PY008 Where CLTCOD = '" + CLTCOD + "' AND  PosDate = '" + PosDate + "' And MSTCOD = '" + MSTCOD + "' ";
                                oRecordSet.DoQuery(sQry);
                            }
                            ProgressBar01.Value += 1;
                            ProgressBar01.Text = ProgressBar01.Value + "/" + oDS_PH_PY008.Rows.Count + "건 삭제중...!";
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                ProgressBar01.Stop();
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                ProgressBar01.Stop();
                PH_PY008_MTX01("A");
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 추가 수정 삭제 가능 여부 조회
        /// </summary>
        /// <returns></returns>
        private string PH_PY008_ModifyYN()
        {
            string today;
            string todaytm;
            string returnValue = string.Empty;
            short errNum = 0;                                        
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);       
            try
            {
                //안강사업장은 제외
                if (oForm.Items.Item("SCLTCOD").Specific.Value != "3")
                {


                    if (oForm.Items.Item("WorkType").Specific.Value.Trim() == "") // 근태구분 등록체크
                    {
                        returnValue = "N";
                        errNum = 3;
                        throw new Exception();
                    }
                    if (oForm.Items.Item("WorkType").Specific.Value.Trim() == "A00" || oForm.Items.Item("WorkType").Specific.Value.Trim() == "D09" || oForm.Items.Item("WorkType").Specific.Value.Trim() == "D10") // 근태구분 등록체크
                    {
                        if (Convert.ToInt32(Convert.ToDouble(oForm.Items.Item("Base").Specific.Value.Substring(0,2))) == 0 ) // 기본+ 특근이 0 이면 등록 안됨.
                        {
                            returnValue = "N";
                            errNum = 4;
                            throw new Exception();
                        }
                    }

                    today = DateTime.Now.ToString("yyyyMMdd");
                    todaytm = DateTime.Now.ToString("HHMM");
                    if (Convert.ToInt32(oForm.Items.Item("PosDate").Specific.Value) < Convert.ToInt32(today))
                    {
                        if (oRspCodeYN == "Y")
                        {
                            returnValue = "Y";
                        }
                        else
                        {
                            returnValue = "N";
                            errNum = 1;
                        }
                    }
                    else if (Convert.ToInt32(oForm.Items.Item("PosDate").Specific.Value) == Convert.ToInt32(today))
                    {
                        if (Convert.ToInt16(todaytm) > 1500)
                        {
                            if (oRspCodeYN == "Y")
                            {
                                returnValue = "Y";
                            }
                            else
                            {
                                returnValue = "N";
                                errNum = 2;
                            }
                        }
                        else
                        {
                            returnValue = "Y";
                        }
                    }

                    else
                    {
                        returnValue = "Y";
                    }
                    
                }
                else
                {
                    returnValue = "Y";
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이전일자는 추가/수정/삭제를 할 수 없습니다. 관리담당으로 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("15시 이후 추가/수정/삭제를 할 수 없습니다. 관리담당으로 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.MessageBox("근태구분을 선택하세요.");
                    oForm.Items.Item("WorkType").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.MessageBox("출/퇴근시간을 입력해주세요");
                    oForm.Items.Item("GetTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            oForm.Freeze(true);
            
            try
            {
                if (pVal.BeforeAction == true)
                {

                    if (pVal.ItemUID == "Btn_ret")
                    {
                        oForm.DataSources.UserDataSources.Item("Chk").Value = "N";

                        oForm.Items.Item("WorkType").Specific.Select(oWorkType, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        PH_PY008_Time_ReSet();

                        PH_PY008_MTX01("A");
                    }

                    if (pVal.ItemUID == "Btn01")
                    {
                        if (PH_PY008_ModifyYN() == "Y")
                        {
                            PH_PY008_SAVE();
                        }
                    }

                    if (pVal.ItemUID == "Btn02")
                    {
                        oForm.DataSources.UserDataSources.Item("Chk").Value = "Y";
                        PH_PY008_MTX01("S"); //근태이상자
                    }

                    if (pVal.ItemUID == "Btn_del")
                    {
                        if (PH_PY008_ModifyYN() == "Y")
                        {
                            PH_PY008_Delete();
                        }
                    }

                    if (pVal.ItemUID == "Btn_Chk")
                    {
                        if (oDS_PH_PY008.Rows.Count >= 0)
                        {
                            for (int i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                            {
                                if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value == "Y")
                                {
                                    oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value = "N";
                                }
                                else
                                {
                                    oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value = "Y";
                                }
                            }

                            oForm.Items.Item("Btn_Chk").Specific.Caption = "선택";
                        }
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
                                    PH_PY008_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY008_FormItemEnabled();
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.ActionSuccess == true)
                                {
                                    PH_PY008_FormItemEnabled();
                                }
                            }
                            break;
                    }
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "SShiftDat")
                        {
                            if (oForm.Items.Item("SShiftDat").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "SGNMUJO")
                        {
                            if (oForm.Items.Item("SGNMUJO").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "SMSTCOD")
                        {
                            if (oForm.Items.Item("SMSTCOD").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "MSTCOD")
                        {
                            if (oForm.Items.Item("MSTCOD").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ItemUID == "ActCode")
                        {
                            if (oForm.Items.Item("ActCode").Specific.Value == "")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                                return;
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i;
            double JanQty = 0;
            string ymd;
            string CLTCOD;
            string MSTCOD;
            string YY;
            short errNum = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        switch (pVal.ItemUID)
                        {
                            case "SCLTCOD":

                                if (oForm.Items.Item("STeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("STeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("STeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("STeamCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("STeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L]";
                                sQry += "             WHERE Code = '1'";
                                sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             AND U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("STeamCode").Specific, "Y");

                                oForm.Items.Item("STeamCode").DisplayDesc = true;

                                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += " WHERE       Code = '1'";
                                sQry += "             And U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                                oForm.Items.Item("TeamCode").DisplayDesc = true;

                                //위해코드
                                if (oForm.Items.Item("DangerCD").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("DangerCD").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("DangerCD").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("DangerCD").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("DangerCD").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += " WHERE       Code = 'P220'";
                                sQry += "             And U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("DangerCD").Specific, "Y");

                                oForm.Items.Item("DangerCD").DisplayDesc = true;
                                break;
                                
                            case "STeamCode":
                                
                                if (oForm.Items.Item("SRspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("SRspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("SRspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("SRspCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("SRspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += " WHERE       Code = '2'";
                                sQry += "             And U_Char1 = '" + oForm.Items.Item("STeamCode").Specific.Value + "'";
                                sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             AND U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("SRspCode").Specific, "Y");

                                oForm.Items.Item("SRspCode").DisplayDesc = true;
                                break;
                                
                            case "TeamCode":
                                
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += " WHERE       Code = '2'";
                                sQry += "             And U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value + "'";
                                sQry += "             AND U_Char2 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                oForm.Items.Item("RspCode").DisplayDesc = true;
                                break;
                                
                            case "SRspCode":
                                    
                                if (oForm.Items.Item("SClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("SClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("SClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }

                                    oForm.Items.Item("SClsCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("SClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += " WHERE       Code = '9'";
                                sQry += "             And U_Char1 = '" + oForm.Items.Item("SRspCode").Specific.Value + "'";
                                sQry += "             AND U_Char3 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("SClsCode").Specific, "Y");

                                oForm.Items.Item("SClsCode").DisplayDesc = true;
                                break;
                                
                            case "RspCode":
                                
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                    oForm.Items.Item("ClsCode").Specific.ValidValues.Add("", "");
                                    oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                sQry = "  SELECT      U_Code,";
                                sQry += "             U_CodeNm";
                                sQry += " FROM        [@PS_HR200L] ";
                                sQry += "             WHERE Code = '9'";
                                sQry += "             And U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
                                sQry += "             AND U_Char3 = '" + oForm.Items.Item("SCLTCOD").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Seq";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");

                                oForm.Items.Item("ClsCode").DisplayDesc = true;
                                break;
                                
                            case "WorkType":
                                
                                switch (oForm.Items.Item("WorkType").Specific.Value.ToString().Trim())
                                {
                                    case "A00":
                                        
                                        //정상근무
                                        oForm.Items.Item("Rotation").Specific.Value = 1;
                                        break;
                                        
                                    case "A01":
                                    case "A02":
                                    case "E02":
                                    case "E03":
                                    case "F01":
                                    case "F02":
                                    case "F03":
                                    case "F04":
                                    case "F05":
                                    case "D11":
                                         
                                        //무단결근, 유계결근, 무급휴일, 휴업, 공상휴직, 병가(휴직), 신병휴직, 정직(유결), 가사휴직
                                        oForm.Items.Item("OffDate").Specific.Value = oForm.Items.Item("GetDate").Specific.Value;
                                        oForm.Items.Item("GetTime").Specific.Value = "0000";
                                        oForm.Items.Item("OffTime").Specific.Value = "0000";

                                        PH_PY008_Time_ReSet();

                                        oForm.Items.Item("Rotation").Specific.Value = 0;
                                        break;
                                        
                                    case "C02":
                                    case "D04":
                                    case "D05":
                                    case "D06":
                                    case "D07":
                                    case "H05":

                                        //훈련, 경조휴가, 하기휴가, 특별휴가, 분만휴가, 조합활동
                                        oForm.Items.Item("GetTime").Specific.Value = "0000";
                                        oForm.Items.Item("OffTime").Specific.Value = "0000";
                                        PH_PY008_Time_ReSet();
                                        oForm.Items.Item("OffDate").Specific.Value = oForm.Items.Item("GetDate").Specific.Value;

                                        oForm.Items.Item("Rotation").Specific.Value = 0;
                                        break;
                                        
                                    case "D02":
                                    case "D09":

                                        //연차/반차 휴가
                                        //연차/반차 휴가 잔여일수 확인
                                        if (oForm.Items.Item("WorkType").Specific.Value.ToString().Trim() == "D02")
                                        {
                                            JanQty = 1;
                                        }
                                        else if (oForm.Items.Item("WorkType").Specific.Value.ToString().Trim() == "D09")
                                        {
                                            JanQty = 0.5;
                                        }

                                        ymd = oForm.Items.Item("PosDate").Specific.Value;
                                        CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;
                                        MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                                        if (MSTCOD.ToString().Trim() != "")
                                        {
                                            sQry = "exec [PH_PY775_01] '" + CLTCOD + "','";
                                            sQry += codeHelpClass.Left(ymd, 4) + "','" + MSTCOD + "'";
                                            oRecordSet.DoQuery(sQry);

                                            if (oRecordSet.Fields.Item("jandd").Value < JanQty)
                                            {
                                                errNum = 1;
                                                oForm.Items.Item("WorkType").Specific.Select("A00", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                throw new Exception();
                                            }
                                            else
                                            {
                                                oForm.Items.Item("GetTime").Specific.Value = "0000";
                                                oForm.Items.Item("OffTime").Specific.Value = "0000";
                                                PH_PY008_Time_ReSet();
                                                oForm.Items.Item("OffDate").Specific.Value = oForm.Items.Item("GetDate").Specific.Value;

                                                oForm.Items.Item("Rotation").Specific.Value = 1;
                                            }
                                        }
                                        else
                                        {
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("사번정보가 없습니다. 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);

                                            oForm.Items.Item("WorkType").Specific.Select("A00", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        break;
                                        
                                    case "D08":
                                    case "D10":
                                        
                                        //근속보전휴가, 근속보전반차(기계사업부)
                                        //근속보전휴가 잔량 확인
                                        ymd = oForm.Items.Item("PosDate").Specific.Value;
                                        CLTCOD = oForm.Items.Item("SCLTCOD").Specific.Value;
                                        MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value;

                                        if (Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("PosDate").Specific.Value, "-")) >= Convert.ToDateTime(ymd.Substring(0, 4) + "-07-01"))
                                        {
                                            YY = codeHelpClass.Left(ymd, 4);
                                        }
                                        else
                                        {
                                            YY = Convert.ToString(Convert.ToInt32(codeHelpClass.Left(ymd, 4)) - 1);
                                        }

                                        sQry = "exec [PH_PY605_99] '" + CLTCOD + "','" + YY + "', '" + ymd + "', '" + MSTCOD + "'";

                                        oRecordSet.DoQuery(sQry);

                                        if (oForm.Items.Item("WorkType").Specific.Value.ToString().Trim() == "D08")
                                        {
                                            JanQty = 1;
                                        }
                                        else if (oForm.Items.Item("WorkType").Specific.Value.ToString().Trim() == "D10")
                                        {
                                            JanQty = 0.5;
                                        }

                                        if (oRecordSet.Fields.Item("Bqty").Value - oRecordSet.Fields.Item("Sqty").Value < JanQty)
                                        {
                                            errNum = 2;
                                            oForm.Items.Item("WorkType").Specific.Select("A00", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                            throw new Exception();
                                        }
                                        else
                                        {
                                            oForm.Items.Item("GetTime").Specific.Value = "0000";
                                            oForm.Items.Item("OffTime").Specific.Value = "0000";
                                            PH_PY008_Time_ReSet();
                                            oForm.Items.Item("OffDate").Specific.Value = oForm.Items.Item("GetDate").Specific.Value;

                                            oForm.Items.Item("Rotation").Specific.Value = 1;
                                        }
                                        break;
                                }

                                break;

                            case "ShiftDat":
                                
                                //근무형태에 따른 근무조 값
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
                                sQry += "             AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.Value + "'";
                                sQry += "             And U_UseYN = 'Y'";
                                sQry += " ORDER BY    U_Code";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "Y");

                                oForm.Items.Item("GNMUJO").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("GNMUJO").DisplayDesc = true;
                                break;
                                
                            case "GNMUJO":
                                
                                //시간reset
                                PH_PY008_Time_ReSet();
                                PH_PY008_ChGNMUJO_Set();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("일반년차잔여일수가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("보전휴가잔여일수가 없습니다. 확인바랍니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
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
            int chkCnt = 0;

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        if (oForm.DataSources.UserDataSources.Item("Chk").Value == "N")
                                        {
                                            if (pVal.ColUID != "Chk")
                                            {
                                                PH_PY008_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                            }
                                        }
                                        else
                                        {
                                            PH_PY008_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                        }
                                        break;
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
                else if (pVal.Before_Action == false)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        
                                        if (oForm.DataSources.UserDataSources.Item("Chk").Value == "N")
                                        {
                                            if (pVal.ColUID == "Chk")
                                            {
                                                //체크수
                                                if (oDS_PH_PY008.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                                                    {
                                                        if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value == "Y")
                                                        {
                                                            chkCnt += 1;
                                                        }
                                                    }
                                                    oForm.Items.Item("Chkcnt").Specific.Value = chkCnt;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            for (int i = 0; i <= oDS_PH_PY008.Rows.Count - 1; i++)
                                            {
                                                if (oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value == "Y")
                                                {
                                                    oDS_PH_PY008.Columns.Item("Chk").Cells.Item(i).Value = "N";
                                                }
                                            }

                                            oDS_PH_PY008.Columns.Item("Chk").Cells.Item(pVal.Row).Value = "Y";
                                            oForm.Items.Item("Chkcnt").Specific.Value = 1;
                                        }
                                        break;
                                }
                            }
                            break;
                    }
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        
                        switch (pVal.ItemUID)
                        {
                            case "SPosDate":
                                
                                oGrid1.DataTable.Clear();
                                oForm.Items.Item("Chkcnt").Specific.Value = 0;
                                break;
                                
                            case "OffDate":
                                
                                oForm.DataSources.UserDataSources.Item("OffTime").Value = "0000";
                                break;
                                
                            case "GetTime":
                                
                                oForm.DataSources.UserDataSources.Item("OffTime").Value = "0000";
                                PH_PY008_Time_ReSet();
                                break;
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "SPosDate":
                                
                                oForm.Freeze(true);
                                PH_PY008_ChDate_Set();
                                PH_PY008_Time_ReSet(); //시간 reset

                                oForm.Items.Item("SShiftDat").Specific.Value = "";
                                oForm.Freeze(false);
                                break;
                                
                            case "PosDate":
                                
                                oForm.Items.Item("SPosDate").Specific.Value = oForm.Items.Item("PosDate").Specific.Value;
                                break;
                                
                            case "GetDate":

                                oForm.DataSources.UserDataSources.Item("SPosDate").Value = oForm.Items.Item("GetDate").Specific.Value;
                                break;

                            case "OffDate":
                                
                                break;

                            case "SShiftDat": //근무형태
                                
                                if (oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() != "")
                                {
                                    sQry = "  SELECT      U_CodeNm";
                                    sQry += " FROM        [@PS_HR200L] ";
                                    sQry += " WHERE       Code = 'P154'";
                                    sQry += "             AND U_Code = '" + oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() + "'";
                                    sQry += " ORDER BY    U_Code";
                                    oRecordSet.DoQuery(sQry);

                                    if (oRecordSet.RecordCount > 0)
                                    {
                                        oForm.Items.Item("ShiftDatNm").Specific.Value = oRecordSet.Fields.Item(0).Value;

                                        sQry = "  select      u_Team1";
                                        sQry += " from        [@PH_PY003B]";
                                        sQry += " where       left(code,1) = '1'";
                                        sQry += "             and u_date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                        oRecordSet.DoQuery(sQry);

                                        oForm.Items.Item("Team1").Specific.Value = oRecordSet.Fields.Item(0).Value;

                                        sQry = "  select      u_Team2";
                                        sQry += " from        [@PH_PY003B]";
                                        sQry += " where       left(code,1) = '1'";
                                        sQry += "             and u_date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                        oRecordSet.DoQuery(sQry);

                                        oForm.Items.Item("Team2").Specific.Value = oRecordSet.Fields.Item(0).Value;

                                        sQry = "  select      u_Team3";
                                        sQry += " from        [@PH_PY003B]";
                                        sQry += " where       left(code,1) = '1'";
                                        sQry +="              and u_date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                        oRecordSet.DoQuery(sQry);

                                        oForm.Items.Item("Team3").Specific.Value = oRecordSet.Fields.Item(0).Value;

                                        sQry = "  SELECT      U_CodeNm";
                                        sQry += " FROM        [@PS_HR200L] ";
                                        sQry += " WHERE       Code = 'P154'";
                                        sQry += "             AND U_Code = '" + oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() + "'";
                                        sQry += " ORDER BY    U_Code";

                                        oRecordSet.DoQuery(sQry);

                                        oForm.Items.Item("SGNMUJO").Specific.Value = "";
                                        oForm.Items.Item("SGNMUJONm").Specific.Value = "";

                                        oForm.Items.Item("ShiftDat").Specific.Select(oRecordSet.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }
                                    else
                                    {
                                        oForm.Items.Item("ShiftDatNm").Specific.Value = "";
                                        PSH_Globals.SBO_Application.StatusBar.SetText("근무형태가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    }
                                }
                                else
                                {
                                    oForm.Items.Item("ShiftDatNm").Specific.Value = "";
                                    oForm.Items.Item("SGNMUJO").Specific.Value = "";
                                    oForm.Items.Item("SGNMUJONm").Specific.Value = "";
                                    oForm.Items.Item("ShiftDat").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                break;

                            case "SGNMUJO": //근무조
                     
                                if (oForm.Items.Item("SGNMUJO").Specific.Value != "")
                                {
                                    sQry = "  SELECT      U_CodeNm ";
                                    sQry += " FROM        [@PS_HR200L] ";
                                    sQry += " WHERE       Code = 'P155' ";
                                    sQry += "             AND U_Code = '" + oForm.Items.Item("SGNMUJO").Specific.Value.ToString().Trim() + "'";
                                    sQry += "             AND U_Char1 = '" + oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() + "'";
                                    sQry += " ORDER BY    U_Code";

                                    oRecordSet.DoQuery(sQry);

                                    if (oForm.Items.Item("SShiftDat").Specific.Value.ToString().Trim() == "4")
                                    {
                                        // WORKTYPE 변경 로직 주 52시간일 경우 휴무일경우 E01 주간/야간일 경우 A00을 입력) (주52)
                                        if (oRecordSet.Fields.Item(0).Value == "1조")
                                        {
                                            sQry = "  Select (case when U_Team1 is null then 'E01' else U_WorkType2 end)";
                                            sQry += " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code ";
                                            sQry += " Where  a.U_CLTCOD = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                                            sQry += "        And b.U_Date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                            oRecordSet01.DoQuery(sQry);
                                        }

                                        if (oRecordSet.Fields.Item(0).Value == "2조")
                                        {
                                            sQry = "  Select (case when U_Team2 is null then 'E01' else U_WorkType2 end)";
                                            sQry += " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code ";
                                            sQry += " Where  a.U_CLTCOD = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                                            sQry += "        And b.U_Date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                            oRecordSet01.DoQuery(sQry);
                                        }

                                        if (oRecordSet.Fields.Item(0).Value == "3조")
                                        {
                                            sQry = "  Select (case when U_Team3 is null then 'E01' else U_WorkType2 end)";
                                            sQry += " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code ";
                                            sQry += " Where  a.U_CLTCOD = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                                            sQry += "        And b.U_Date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                            oRecordSet01.DoQuery(sQry);
                                        }

                                        if (oRecordSet01.RecordCount == 0)
                                        {
                                            PSH_Globals.SBO_Application.StatusBar.SetText("근태월력의 근무일 등록을 하지않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            oWorkType = oRecordSet01.Fields.Item(0).Value;
                                            oForm.Items.Item("WorkType").Specific.Select(oRecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                    }
                                    else
                                    {
                                        sQry = "  SELECT  T1.U_WorkType";
                                        sQry += " FROM    [@PH_PY003A] AS T0 ";
                                        sQry += "         INNER JOIN ";
                                        sQry += "         [@PH_PY003B] AS T1 ";
                                        sQry += "             ON T0.Code = T1.Code ";
                                        sQry += " Where   T0.U_CLTCOD = '" + oForm.Items.Item("SCLTCOD").Specific.Value.ToString().Trim() + "'";
                                        sQry += "         AND T1.U_Date = '" + oForm.Items.Item("SPosDate").Specific.Value.ToString().Trim() + "'";

                                        oRecordSet01.DoQuery(sQry);

                                        oForm.Items.Item("WorkType").Specific.Select(oRecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }

                                    if (oRecordSet.RecordCount > 0)
                                    {   
                                        oForm.Items.Item("SGNMUJONm").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                        oForm.Items.Item("GNMUJO").Specific.Select(oRecordSet.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                        PH_PY008_ChGNMUJO_Set();
                                    }
                                    else
                                    {
                                        oForm.Items.Item("SGNMUJONm").Specific.Value = "";
                                        PSH_Globals.SBO_Application.StatusBar.SetText("근무형태가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    }
                                }
                                break;

                            case "GetTime":
                                    break;
                                
                            case "OffTime":

                                if (oForm.Items.Item("OffTime").Specific.Value != "0000")
                                {
                                    PH_PY008_Time_Calc_Main(oForm.Items.Item("OffTime").Specific.Value);
                                }
                                break;

                            case "ActCode":

                                sQry = "  select  b.U_CodeNm,";
                                sQry += "         b.U_Char1";
                                sQry += " from    [@PS_HR200H] a";
                                sQry += "         Inner Join";
                                sQry += "         [@PS_HR200L] b";
                                sQry += "             On a.Code = b.Code";
                                sQry += "             And a.Code = 'P127'";
                                sQry += "             And b.U_Code = '" + oForm.Items.Item("ActCode").Specific.Value + "'";

                                oRecordSet.DoQuery(sQry);
                                
                                if (oRecordSet.RecordCount > 0)
                                {
                                    oForm.Items.Item("ActText").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                    oForm.Items.Item("DangerCD").Specific.Select(oRecordSet.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    oForm.Items.Item("ActText").Specific.Value = "";
                                    oForm.Items.Item("DangerCD").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                break;
                                
                            case "SMSTCOD":

                                sQry = "  SELECT      U_FullName";
                                sQry += " from        [@PH_PY001A]";
                                sQry += " Where       Code = '" + oForm.Items.Item("SMSTCOD").Specific.Value + "'";

                                oRecordSet.DoQuery(sQry);

                                if (oRecordSet.RecordCount > 0)
                                {
                                    oForm.Items.Item("SFullName").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                }
                                else
                                {
                                    oForm.Items.Item("SFullName").Specific.Value = "";
                                }
                                break;
                                
                            case "MSTCOD":

                                sQry = "  SELECT      U_FullName,";
                                sQry += "             U_TeamCode,";
                                sQry += "             U_RspCode,";
                                sQry += "             U_ShiftDat,";
                                sQry += "             U_GNMUJO";
                                sQry += " from        [@PH_PY001A]";
                                sQry += " Where       Code = '" + oForm.Items.Item("MSTCOD").Specific.Value + "'";

                                oRecordSet.DoQuery(sQry);

                                if (oRecordSet.RecordCount > 0)
                                {
                                    oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item(0).Value;
                                    oForm.Items.Item("TeamCode").Specific.Select(oRecordSet.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.Items.Item("RspCode").Specific.Select(oRecordSet.Fields.Item(2).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.Items.Item("ShiftDat").Specific.Select(oRecordSet.Fields.Item(3).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm.Items.Item("GNMUJO").Specific.Select(oRecordSet.Fields.Item(4).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    oForm.Items.Item("FullName").Specific.Value = "";
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY008);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            PH_PY008_FormItemEnabled();
                            break;
                            
                        case "1284":
                            break;
                            
                        case "1286":
                            break;
                            
                        case "1281": //문서찾기
               
                            PH_PY008_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;

                        case "1282": //문서추가

                            PH_PY008_FormItemEnabled();
                            break;
                 
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":

                            PH_PY008_FormItemEnabled();
                            break;

                        case "1293": // 행삭제
                            break;
                    }
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
