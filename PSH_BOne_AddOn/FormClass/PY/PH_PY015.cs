using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 연차이월등록
    /// </summary>
    internal class PH_PY015 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Matrix oMat01;

        private SAPbouiCOM.DBDataSource oDS_PH_PY015B; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private int oLast_Mode;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            string strXml = string.Empty;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY015.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY015_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY015");

                strXml = oXmlDoc.xml.ToString();
                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry"

                oForm.Freeze(true);
                PH_PY015_CreateItems();
                //PH_PY015_ComboBox_Setting();
                //PH_PY015_CF_ChooseFromList();
                PH_PY015_EnableMenus();
                PH_PY015_SetDocument(oFromDocEntry01);
                //PH_PY015_FormResize();
                PH_PY015_LoadCaption();
                PH_PY015_FormReset();
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
                //oForm.ActiveItem = "Date";
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY015_CreateItems()
        {
            string oQuery01 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY015B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                //메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");

                //부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

                //부서명
                oForm.DataSources.UserDataSources.Add("TeamName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("TeamName").Specific.DataBind.SetBound(true, "", "TeamName");

                //담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

                //담당명
                oForm.DataSources.UserDataSources.Add("RspName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("RspName").Specific.DataBind.SetBound(true, "", "RspName");

                //반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                //반명
                oForm.DataSources.UserDataSources.Add("ClsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ClsName").Specific.DataBind.SetBound(true, "", "ClsName");

                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");

                //사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                //성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY015_EnableMenus()
        {
            try
            {
                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), false); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1285"), false); // 복원
                oForm.EnableMenu(("1284"), false); // 취소
                oForm.EnableMenu(("1293"), false); // 행삭제
                oForm.EnableMenu(("1281"), false);
                oForm.EnableMenu(("1282"), true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY015_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY015_FormItemEnabled();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY015_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }

                //매트릭스 컬럼(사업장코드, 기준년도, 사번) 숨김
                oMat01.Columns.Item("CLTCOD").Visible = false;
                oMat01.Columns.Item("StdYear").Visible = false;
                oMat01.Columns.Item("MSTCOD").Visible = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PH_PY015_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "저장";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "저장";
                }

                oForm.Freeze(false);
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
        /// 화면의 크기 재조정
        /// </summary>
        private void PH_PY015_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PH_PY015_FormReset()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                //헤더 초기화
                oForm.DataSources.UserDataSources.Item("CLTCOD").Value = dataHelpClass.User_BPLID(); //사업장
                oForm.DataSources.UserDataSources.Item("TeamCode").Value = ""; //부서
                oForm.DataSources.UserDataSources.Item("TeamName").Value = ""; //부서명
                oForm.DataSources.UserDataSources.Item("RspCode").Value = ""; //담당
                oForm.DataSources.UserDataSources.Item("RspName").Value = ""; //담당명
                oForm.DataSources.UserDataSources.Item("ClsCode").Value = ""; //반
                oForm.DataSources.UserDataSources.Item("ClsName").Value = ""; //반명
                oForm.DataSources.UserDataSources.Item("StdYear").Value = DateTime.Now.ToString("yyyy"); //기준년도 //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")
                oForm.DataSources.UserDataSources.Item("MSTCOD").Value = ""; //사번
                oForm.DataSources.UserDataSources.Item("FullName").Value = ""; //성명

                //라인 초기화
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();
                //PH_PY015_Add_MatrixRow(0, True)
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
        /// 매트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted">기본 RowIserted false</param>
        private void PH_PY015_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false) //행추가여부
                {
                    oDS_PH_PY015B.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PH_PY015B.Offset = oRow;
                oDS_PH_PY015B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PH_PY015B.SetValue("U_ColReg01", oRow, "Y");

                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY015_MTX01()
        {
            short i = 0;
            short errNum = 0;
            string sQry = string.Empty;

            string CLTCOD = string.Empty; //사업장
            string TeamCode = string.Empty; //부서
            string RspCode = string.Empty; //담당
            string ClsCode = string.Empty; //반
            string StdYear = string.Empty; //기준년도
            string MSTCOD = string.Empty; //사번

            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim(); //사업장
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim(); //부서
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim(); //담당
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim(); //반
                StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim(); //기준년도
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() ; //사번

                sQry = "      EXEC [PH_PY015_01] '";
                sQry = sQry + CLTCOD + "','"; //사업장
                sQry = sQry + TeamCode + "','"; //부서
                sQry = sQry + RspCode + "','"; //담당
                sQry = sQry + ClsCode + "','"; //반
                sQry = sQry + StdYear + "','"; //기준년도
                sQry = sQry + MSTCOD + "'"; //사번

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY015B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errNum = 1;

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                    PH_PY015_Add_MatrixRow(0, true);
                    PH_PY015_LoadCaption();

                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY015B.Size)
                    {
                        oDS_PH_PY015B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY015B.Offset = i;

                    oDS_PH_PY015B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY015B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("Check").Value.ToString().Trim()); //선택
                    oDS_PH_PY015B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim()); //부서
                    oDS_PH_PY015B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("RspName").Value.ToString().Trim()); //담당
                    oDS_PH_PY015B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("ClsName").Value.ToString().Trim()); //반
                    oDS_PH_PY015B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim()); //성명
                    oDS_PH_PY015B.SetValue("U_ColDt01", i, oRecordSet01.Fields.Item("ipsail").Value.ToString("yyyyMMdd")); //입사일자
                    oDS_PH_PY015B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("sojungil").Value.ToString().Trim()); //근로의무일
                    oDS_PH_PY015B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("bicnt").Value.ToString().Trim()); //결근일
                    oDS_PH_PY015B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("bRate").Value.ToString().Trim()); //출근율
                    oDS_PH_PY015B.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("yycha").Value.ToString().Trim()); //연차발생
                    oDS_PH_PY015B.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("gunsok").Value.ToString().Trim()); //근속가산
                    oDS_PH_PY015B.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("iwol").Value.ToString().Trim()); //전년이월
                    oDS_PH_PY015B.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("Tot").Value.ToString().Trim()); //계
                    oDS_PH_PY015B.SetValue("U_ColQty05", i, oRecordSet01.Fields.Item("useyy").Value.ToString().Trim()); //사용일수
                    oDS_PH_PY015B.SetValue("U_ColQty06", i, oRecordSet01.Fields.Item("jandd").Value.ToString().Trim()); //정산대상
                    oDS_PH_PY015B.SetValue("U_ColQty07", i, oRecordSet01.Fields.Item("savedd").Value.ToString().Trim()); //적치
                    oDS_PH_PY015B.SetValue("U_ColQty08", i, oRecordSet01.Fields.Item("paydd").Value.ToString().Trim()); //임금대치
                    oDS_PH_PY015B.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim()); //사업장
                    oDS_PH_PY015B.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("StdYear").Value.ToString().Trim()); //기준년도
                    oDS_PH_PY015B.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //사번

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                //PH_PY015_Add_MatrixRow(oMat01.VisualRowCount, false);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                ProgBar01.Stop();
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgBar01.Stop();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// 적치수량 삭제
        /// </summary>
        private void PH_PY015_DeleteData()
        {
            short loopCount = 0;
            string sQry = string.Empty;
            //short errNum = 0;
            string CLTCOD = string.Empty;
            string StdYear = string.Empty;
            string MSTCOD = string.Empty;

            SAPbouiCOM.ProgressBar ProgBar01 = null;
            ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("삭제 중...", 100, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY015B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {

                        CLTCOD = oDS_PH_PY015B.GetValue("U_ColReg18", loopCount).ToString().Trim(); //사업장코드
                        StdYear = oDS_PH_PY015B.GetValue("U_ColReg19", loopCount).ToString().Trim(); //기준년도
                        MSTCOD = oDS_PH_PY015B.GetValue("U_ColReg20", loopCount).ToString().Trim(); //사번

                        sQry = "      EXEC [PH_PY015_04] '";
                        sQry = sQry + CLTCOD + "','"; //사업장코드
                        sQry = sQry + StdYear + "','"; //기준년도
                        sQry = sQry + MSTCOD + "'"; //사번

                        oRecordSet01.DoQuery(sQry);
                    }
                }
                
                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch(Exception ex)
            {
                ProgBar01.Stop();

                //if (errNum == 1)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText("삭제대상이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                //}
                //else if (errNum == 2)
                //{
                //    PSH_Globals.SBO_Application.StatusBar.SetText(loopCount + 1 + "행은 실적이 등록되어 삭제가 불가능합니다. " + loopCount + 1 + "행 이후 삭제는 모두 취소됩니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                //}
                //else
                //{
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //}
            }
            finally
            {
                ProgBar01.Value = 100;
                ProgBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
            }
        }

        /// <summary>
        /// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PH_PY015_AddData()
        {
            bool functionReturnValue = false;
            
            short i = 0;
            string sQry = string.Empty;
            short errNum = 0;

            string CLTCOD = string.Empty; //사업장코드
            string StdYear = string.Empty; //기준년도
            string MSTCOD = string.Empty; //사번
            string FullName = string.Empty; //성명
            float SaveDCnt = 0; //적치(이월) 수량
            float PayDCnt = 0; //임금대치 수량
            float UseDCnt = 0; //사용 수량
            string UserSign = string.Empty; //등록자ID

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("저장 중...", oMat01.VisualRowCount - 1, false);

                UserSign = PSH_Globals.oCompany.UserSignature.ToString();

                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (oDS_PH_PY015B.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        CLTCOD = oDS_PH_PY015B.GetValue("U_ColReg18", i).ToString().Trim(); //사업장코드
                        StdYear = oDS_PH_PY015B.GetValue("U_ColReg19", i).ToString().Trim(); //기준년도
                        MSTCOD = oDS_PH_PY015B.GetValue("U_ColReg20", i).ToString().Trim(); //사번
                        FullName = oDS_PH_PY015B.GetValue("U_ColReg05", i).ToString().Trim(); //성명
                        SaveDCnt = Convert.ToSingle(oDS_PH_PY015B.GetValue("U_ColQty07", i).ToString().Trim()); //적치(이월) 수량
                        PayDCnt = Convert.ToSingle(oDS_PH_PY015B.GetValue("U_ColQty08", i).ToString().Trim()); //임금대치 수량
                        UseDCnt = Convert.ToSingle(oDS_PH_PY015B.GetValue("U_ColQty05", i).ToString().Trim()); //사용 수량

                        sQry = "      EXEC [PH_PY015_02] '";
                        sQry = sQry + CLTCOD + "','"; //사업장코드
                        sQry = sQry + StdYear + "','";//기준년도
                        sQry = sQry + MSTCOD + "','"; //사번
                        sQry = sQry + FullName + "','"; //성명
                        sQry = sQry + SaveDCnt + "','"; //적치(이월) 수량
                        sQry = sQry + PayDCnt + "','"; //임금대치 수량
                        sQry = sQry + UseDCnt + "','"; //사용수량
                        sQry = sQry + UserSign + "'"; //등록자ID

                        if (SaveDCnt == 0 && PayDCnt == 0)
                        {
                            string stringSpace = new string(' ', 10);

                            if (PSH_Globals.SBO_Application.MessageBox(FullName + " 사원의 적치수량, 임금대치수량이 0입니다." + stringSpace + "당해년도 입사자가 아니면 필수로 등록하여야합니다." + stringSpace + "계속 등록하시겠습니까?", 1, "예", "아니오") == 1)
                            {
                                RecordSet01.DoQuery(sQry);
                            }
                            else
                            {
                                errNum = 1;
                                throw new Exception();
                            }
                        }
                        else
                        {
                            RecordSet01.DoQuery(sQry);
                        }

                        ProgBar01.Value = ProgBar01.Value + 1;
                        ProgBar01.Text = ProgBar01.Value + "/" + (oMat01.VisualRowCount - 1) + "건 저장중...";
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                ProgBar01.Stop();
                functionReturnValue = false;

                if (errNum == 1)
                { 
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력이 취소되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgBar01.Value = 100;
                ProgBar01.Stop();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY015_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            short errNum = 0;
            
            float SaveDCnt = 0; //적치수량
            float JanDCnt = 0; //정산대상 수량

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "Mat01":

                        oMat01.FlushToDataSource();

                        if (oCol == "savedd")
                        {
                            SaveDCnt = Convert.ToSingle(oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.Value);
                            JanDCnt = Convert.ToSingle(oMat01.Columns.Item("jandd").Cells.Item(oRow).Specific.Value);

                            if (SaveDCnt > JanDCnt)
                            {
                                errNum = 1;
                                throw new Exception();
                            }
                            else
                            {
                                oDS_PH_PY015B.SetValue("U_ColQty08", oRow - 1, (JanDCnt - SaveDCnt).ToString());
                            }
                        }

                        #region 백업
                        //            If oCol = "MSTCOD" Then

                        //                Call oDS_PH_PY015B.setValue("U_ColReg04", oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE) '사번
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg03", oRow - 1, oForm.Items("StdYear").Specific.VALUE) '기준년도
                        //
                        //                '대상자의 인사마스터에서 소속 조회
                        //                sQry = "        SELECT  T0.U_TeamCode AS [TeamCode], " '부서코드
                        //                sQry = sQry & "         T1.U_CodeNm AS [TeamName], " '부서명
                        //                sQry = sQry & "         T0.U_RspCode AS [RspCode], " '담당코드
                        //                sQry = sQry & "         T2.U_CodeNm AS [RspName], " '담당명
                        //                sQry = sQry & "         T0.U_ClsCode AS [ClsCode], " '반코드
                        //                sQry = sQry & "         T3.U_CodeNm AS [ClsName], " '반명
                        //                sQry = sQry & "         T0.U_FullName AS [FullName], " '성명
                        //                sQry = sQry & "         T4.U_CodeNm AS [JIGNAM]," '직급
                        //                sQry = sQry & "         T0.U_CLTCOD AS [CLTCOD]" '소속사업장
                        //                sQry = sQry & " FROM    [@PH_PY001A] AS T0 "
                        //                sQry = sQry & "         LEFT JOIN"
                        //                sQry = sQry & "         [@PS_HR200L] AS T1"
                        //                sQry = sQry & "             ON T0.U_TeamCode = T1.U_Code"
                        //                sQry = sQry & "             AND T1.Code = '1'"
                        //                sQry = sQry & "         LEFT JOIN"
                        //                sQry = sQry & "         [@PS_HR200L] AS T2"
                        //                sQry = sQry & "             ON T0.U_RspCode = T2.U_Code"
                        //                sQry = sQry & "             AND T2.Code = '2'"
                        //                sQry = sQry & "         LEFT JOIN"
                        //                sQry = sQry & "         [@PS_HR200L] AS T3"
                        //                sQry = sQry & "             ON T0.U_ClsCode = T3.U_Code"
                        //                sQry = sQry & "             AND T3.Code = '9'"
                        //                sQry = sQry & "         LEFT JOIN"
                        //                sQry = sQry & "         [@PS_HR200L] AS T4"
                        //                sQry = sQry & "             ON T0.U_JIGCOD = T4.U_Code"
                        //                sQry = sQry & "             AND T4.Code = 'P129'"
                        //                sQry = sQry & " WHERE   T0.Code = '" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'"
                        //
                        //                Call oRecordSet01.DoQuery(sQry)
                        //
                        //                TeamCode = oRecordSet01.Fields("TeamCode").VALUE
                        //                TeamName = oRecordSet01.Fields("TeamName").VALUE
                        //                RspCode = oRecordSet01.Fields("RspCode").VALUE
                        //                RspName = oRecordSet01.Fields("RspName").VALUE
                        //                ClsCode = oRecordSet01.Fields("ClsCode").VALUE
                        //                ClsName = oRecordSet01.Fields("ClsName").VALUE
                        //                FullName = IIf(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "9999999", "전사원", oRecordSet01.Fields("FullName").VALUE)
                        //                JIGNAM = oRecordSet01.Fields("JIGNAM").VALUE
                        //                CLTCOD = oRecordSet01.Fields("CLTCOD").VALUE
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg06", oRow - 1, TeamCode) '부서코드
                        //                Call oDS_PH_PY015B.setValue("U_ColReg07", oRow - 1, TeamName) '부서명
                        //                Call oDS_PH_PY015B.setValue("U_ColReg08", oRow - 1, RspCode) '담당코드
                        //                Call oDS_PH_PY015B.setValue("U_ColReg09", oRow - 1, RspName) '담당명
                        //                Call oDS_PH_PY015B.setValue("U_ColReg10", oRow - 1, ClsCode) '반코드
                        //                Call oDS_PH_PY015B.setValue("U_ColReg11", oRow - 1, ClsName) '반명
                        //                Call oDS_PH_PY015B.setValue("U_ColReg05", oRow - 1, FullName) '성명
                        //                Call oDS_PH_PY015B.setValue("U_ColReg12", oRow - 1, JIGNAM) '직급
                        //                Call oDS_PH_PY015B.setValue("U_ColReg26", oRow - 1, CLTCOD) '사업장
                        //
                        //                '행 추가
                        //                If oMat01.RowCount = oRow And Trim(oDS_PH_PY015B.GetValue("U_ColReg04", oRow - 1)) <> "" Then
                        //                    Call PH_PY015_Add_MatrixRow(oRow)
                        //                End If
                        //
                        //                Call oMat01.Columns("JIGNAM").Cells(oRow).CLICK(ct_Regular)
                        //
                        //            ElseIf oCol = "TeamCode" Then
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg07", oRow - 1, IIf(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE = "9999", "전부서", MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '1'"))) '부서
                        //
                        //            ElseIf oCol = "RspCode" Then
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg09", oRow - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '2'")) '담당
                        //
                        //            ElseIf oCol = "ClsCode" Then
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg11", oRow - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'", " AND Code = '9'")) '반
                        //
                        //            ElseIf oCol = "Class" Then
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColReg14", oRow - 1, MDC_GetData.Get_ReData("name", "edType", "OHED", "'" & oMat01.Columns(oCol).Cells(oRow).Specific.VALUE & "'")) '교육구분
                        //
                        //            ElseIf oCol = "EduFrDt" Then '교육시작일
                        //
                        //                EduMonth = Mid(oMat01.Columns(oCol).Cells(oRow).Specific.VALUE, 5, 2)
                        //
                        //                If Val(EduMonth) >= 1 And Val(EduMonth) <= 3 Then
                        //
                        //                    EduQuarter = "1"
                        //                    EduHarf = "상"
                        //
                        //                ElseIf Val(EduMonth) >= 4 And Val(EduMonth) <= 6 Then
                        //
                        //                    EduQuarter = "2"
                        //                    EduHarf = "상"
                        //
                        //                ElseIf Val(EduMonth) >= 7 And Val(EduMonth) <= 9 Then
                        //
                        //                    EduQuarter = "3"
                        //                    EduHarf = "하"
                        //
                        //                ElseIf Val(EduMonth) >= 10 And Val(EduMonth) <= 12 Then
                        //
                        //                    EduQuarter = "4"
                        //                    EduHarf = "하"
                        //
                        //                End If
                        //
                        //                Call oDS_PH_PY015B.setValue("U_ColSum01", oRow - 1, Val(EduQuarter)) '분기
                        //                Call oDS_PH_PY015B.setValue("U_ColReg18", oRow - 1, EduHarf) '반기
                        //                Call oDS_PH_PY015B.setValue("U_ColSum02", oRow - 1, Val(EduMonth)) '월
                        //
                        //                'Call oMat01.Columns("EduToDt").Cells(oRow).CLICK(ct_Regular)
                        //
                        //            End If
                        //
                        #endregion

                        oMat01.LoadFromDataSource();
                        //강제 포커스 이동_S
                        oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //강제 포커스 이동_E
                        oMat01.AutoResizeColumns();
                        break;

                    case "MSTCOD": //성명

                        oForm.Items.Item("FullName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'", "");
                        break;

                    case "TeamCode": //부서

                        oForm.Items.Item("TeamName").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code",  "[@PS_HR200L]", "'" + oForm.Items.Item("TeamCode").Specific.VALUE + "'", "AND Code = '1'");
                        break;

                    case "RspCode": //담당

                        oForm.Items.Item("RspName").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]",  "'" + oForm.Items.Item("RspCode").Specific.VALUE + "'", "AND Code = '2'");
                        break;

                    case "ClsCode": //반

                        oForm.Items.Item("ClsName").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("ClsCode").Specific.VALUE + "'", "AND Code = '9'");
                        break;
                }
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("적치(이월) 수량이 정산대상보다 많습니다. 확인하십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Matrix 체크박스 전체 선택
        /// </summary>
        private void PH_PY015_CheckAll()
        {
            string CheckType = string.Empty;
            short loopCount = 0;

            try
            {
                oForm.Freeze(true);
                CheckType = "Y";

                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PH_PY015B.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    oDS_PH_PY015B.Offset = loopCount;
                    if (CheckType == "N")
                    {
                        oDS_PH_PY015B.SetValue("U_ColReg01", loopCount, "Y");
                    }
                    else
                    {
                        oDS_PH_PY015B.SetValue("U_ColReg01", loopCount, "N");
                    }
                }

                oMat01.LoadFromDataSource();
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "PH_PY015")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }

                    //추가/확인 버튼클릭
                    if (pVal.ItemUID == "BtnAdd")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY015_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY015_LoadCaption();
                            PH_PY015_MTX01();

                            oLast_Mode = Convert.ToInt16(oForm.Mode);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            PH_PY015_FormReset();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY015_LoadCaption();
                            PH_PY015_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        
                        PH_PY015_LoadCaption();
                        PH_PY015_MTX01();
                    }
                    else if (pVal.ItemUID == "BtnDelete") //삭제
                    {
                        if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
                        {
                            PH_PY015_DeleteData();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            
                            PH_PY015_LoadCaption();
                            PH_PY015_MTX01();
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSel") //전체선택
                    {
                        PH_PY015_CheckAll();
                    }
                    else if (pVal.ItemUID == "BtnPrint")
                    {
                        //PH_PY015_Print_Report01();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY015")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "TeamCode", ""); //Header-부서
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "RspCode", ""); //Header-담당
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ClsCode", ""); //Header-반
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MSTCOD", ""); //Header-사번
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else
                    {
                        PH_PY015_FlushToItemValue(pVal.ItemUID, 0, "");
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "savedd")
                            {
                                PH_PY015_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            oMat01.AutoResizeColumns();
                        }
                        else
                        {
                            PH_PY015_FlushToItemValue(pVal.ItemUID, 0, "");
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
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
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
                    PH_PY015_FormItemEnabled();
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

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY015B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PH_PY015_FormResize();
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
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                            PH_PY015_FormReset();
                            PH_PY015_LoadCaption();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기

                            //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            oForm.Freeze(true);
                            PH_PY015_Add_MatrixRow(oMat01.VisualRowCount, false);
                            oForm.Freeze(false);
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기

                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PH_PY015B.RemoveRecord(oDS_PH_PY015B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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


        #region Raise_FormItemEvent
        //		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //					////8
        //					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //					////10
        //					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //					////18
        //					break;
        //				////et_FORM_ACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //					////19
        //					break;
        //				////et_FORM_DEACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //					////20
        //					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //					////27
        //					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;
        //				////et_LOST_FOCUS
        //				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //					////17
        //					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //			}
        //			return;
        //			Raise_FormItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormMenuEvent
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {

        //			}
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {

        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}

        //			return;
        //			Raise_RightClickEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Item EVENTS
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string ErrorRecordMstNam = null;
        //			//적치, 임금대치 둘다 0인 행의 사원 이름 저장

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {


        //			}

        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				//        If (pval.ItemUID = "ItemCode") Then
        //				//            Dim oDataTable01 As SAPbouiCOM.DataTable
        //				//            Set oDataTable01 = pval.SelectedObjects
        //				//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
        //				//            Set oDataTable01 = Nothing
        //				//        End If
        //				//        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
        //				//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY015A", "U_CardCode,U_CardName")
        //				//        End If
        //			}
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			if ((oLastColRow01 > 0)) {
        //				if (pval.BeforeAction == true) {
        //					//            If (PH_PY015_Validate("행삭제") = False) Then
        //					//                BubbleEvent = False
        //					//                Exit Sub
        //					//            End If
        //					////행삭제전 행삭제가능여부검사
        //				} else if (pval.BeforeAction == false) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					oDS_PH_PY015B.RemoveRecord(oDS_PH_PY015B.Size - 1);
        //					oMat01.LoadFromDataSource();
        //					if (oMat01.RowCount == 0) {
        //						PH_PY015_Add_MatrixRow(0);
        //					} else {
        //						if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY015B.GetValue("U_CntcCode", oMat01.RowCount - 1)))) {
        //							PH_PY015_Add_MatrixRow(oMat01.RowCount);
        //						}
        //					}
        //				}
        //			}
        //			return;
        //			Raise_EVENT_ROW_DELETE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region 백업소스코드
        ///// <summary>
        ///// 콤보박스 SET (주석처리)
        ///// </summary>
        //private void PH_PY015_ComboBox_Setting()
        //{
        //    string sQry = string.Empty;
        //    SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    try
        //    {
        //        oForm.Freeze(true);

        //        //    '재직여부
        //        //    Call oForm.Items("Status").Specific.ValidValues.Add("1", "재직자")
        //        //    Call oForm.Items("Status").Specific.ValidValues.Add("2", "퇴직자포함")
        //        //    Call oForm.Items("Status").Specific.Select(0, psk_Index)
        //        //
        //        //    '기준월
        //        //    sQry = "        SELECT      U_Code AS [Code],"
        //        //    sQry = sQry & "             U_CodeNm AS [Name]"
        //        //    sQry = sQry & " FROM        [@PS_HR200L]"
        //        //    sQry = sQry & " WHERE       Code = '4'"
        //        //    sQry = sQry & "             AND U_UseYN = 'Y'"
        //        //    sQry = sQry & " ORDER BY    U_Seq"
        //        //
        //        //    Call oForm.Items("StdMt").Specific.ValidValues.Add("%", "전체")
        //        //    Call MDC_SetMod.Set_ComboList(oForm.Items("StdMt").Specific, sQry, "", False, False)
        //        //    Call oForm.Items("StdMt").Specific.Select(0, psk_Index)
        //        //
        //        //    '계획구분
        //        //    sQry = "        SELECT      U_Code AS [Code],"
        //        //    sQry = sQry & "             U_CodeNm AS [Name]"
        //        //    sQry = sQry & " FROM        [@PS_HR200L]"
        //        //    sQry = sQry & " WHERE       Code = 'P239'"
        //        //    sQry = sQry & "             AND U_UseYN = 'Y'"
        //        //    sQry = sQry & " ORDER BY    U_Seq"
        //        //
        //        //    Call oForm.Items("PlnCls").Specific.ValidValues.Add("%", "전체")
        //        //    Call MDC_SetMod.Set_ComboList(oForm.Items("PlnCls").Specific, sQry, "", False, False)
        //        //    Call oForm.Items("PlnCls").Specific.Select(0, psk_Index)
        //        //
        //        //    '상태
        //        //    Call oForm.Items("DocStatus").Specific.ValidValues.Add("%", "전체")
        //        //    Call oForm.Items("DocStatus").Specific.ValidValues.Add("O", "계획")
        //        //    Call oForm.Items("DocStatus").Specific.ValidValues.Add("P", "완료")
        //        //    Call oForm.Items("DocStatus").Specific.ValidValues.Add("C", "취소")
        //        //    Call oForm.Items("DocStatus").Specific.Select(0, psk_Index)
        //        //
        //        //    '//////////매트릭스//////////_S
        //        //    '상태
        //        //    Call oMat01.Columns("DocStatus").ValidValues.Add("O", "계획")
        //        //    Call oMat01.Columns("DocStatus").ValidValues.Add("P", "완료")
        //        //    Call oMat01.Columns("DocStatus").ValidValues.Add("C", "취소")
        //        //
        //        //    '계획구분
        //        //    sQry = "        SELECT      U_Code AS [Code],"
        //        //    sQry = sQry & "             U_CodeNm AS [Name]"
        //        //    sQry = sQry & " FROM        [@PS_HR200L]"
        //        //    sQry = sQry & " WHERE       Code = 'P239'"
        //        //    sQry = sQry & "             AND U_UseYN = 'Y'"
        //        //    sQry = sQry & " ORDER BY    U_Seq"
        //        //    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PlnCls"), sQry)
        //        //
        //        //    '이수(수료)증
        //        //    Call oMat01.Columns("Certi").ValidValues.Add("O", "미제출")
        //        //    Call oMat01.Columns("Certi").ValidValues.Add("C", "제출")
        //        //
        //        //    '보고서
        //        //    Call oMat01.Columns("Report").ValidValues.Add("O", "미제출")
        //        //    Call oMat01.Columns("Report").ValidValues.Add("C", "제출")
        //        //
        //        //    '사업장
        //        //    sQry = "        SELECT      BPLId AS [BPLId],"
        //        //    sQry = sQry & "             BPLName AS [BPLName]"
        //        //    sQry = sQry & " FROM        OBPL"
        //        //
        //        //    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), sQry)
        //        //    '//////////매트릭스//////////_E
        //    }
        //    catch(Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
        //    }
        //}


        //public void PH_PY015_CF_ChooseFromList()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement


        //    return;
        //PH_PY015_CF_ChooseFromList_Error:
        //    MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY015_CF_ChooseFromList_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}



        ///// <summary>
        ///// 필수입력사항 체크
        ///// </summary>
        ///// <returns></returns>
        //private bool PH_PY015_HeaderSpaceLineDel()
        //{
        //    bool functionReturnValue = false;
        //    short errNum = 0;

        //    try
        //    {
        //        switch (true)
        //        {
        //            case string.IsNullOrEmpty(oForm.Items.Item("DestNo1").Specific.Value): //출장번호1
        //                errNum = 1;
        //                throw new Exception();
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("DestNo2").Specific.VALUE)): //출장번호2
        //                ErrNum = 2;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE)): //사원번호
        //                ErrNum = 3;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrDate").Specific.VALUE)): //시작일자
        //                ErrNum = 4;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("FrTime").Specific.VALUE)): //시작시각
        //                ErrNum = 5;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToDate").Specific.VALUE)): //종료일자
        //                ErrNum = 6;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //            case string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("ToTime").Specific.VALUE)): //종료시각
        //                ErrNum = 7;
        //                goto PH_PY015_HeaderSpaceLineDel_Error;
        //                break;
        //        }

        //        functionReturnValue = true;
        //    }
        //    catch(Exception ex)
        //    {
        //        functionReturnValue = false;

        //        if (ErrNum == 1)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "출장번호1은 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("DestNo1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 2)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "출장번호2는 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("DestNo2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 3)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "사원번호는 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 4)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "시작일자는 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 5)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "시작시각은 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 6)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "종료일자는 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        else if (ErrNum == 7)
        //        {
        //            MDC_Com.MDC_GF_Message(ref "종료시각은 필수사항입니다. 확인하세요.", ref "E");
        //            oForm.Items.Item("FrTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }

        //    }

        //    return functionReturnValue;
        //}

        ///// 메트릭스 필수 사항 check
        //		private bool PH_PY015_MatrixSpaceLineDel()
        //		{
        //			bool functionReturnValue = false;
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			int i = 0;
        //			short ErrNum = 0;
        //			ErrNum = 0;

        //			oMat01.FlushToDataSource();
        //			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //				if (Strings.Trim(oDS_PH_PY015B.GetValue("U_ColReg01", i)) == "Y") {

        //					//적치, 임금대치 둘다 0인행 체크
        //					if (Convert.ToDouble(Strings.Trim(oDS_PH_PY015B.GetValue("U_ColReg07", i))) == 0 & Convert.ToDouble(Strings.Trim(oDS_PH_PY015B.GetValue("U_ColReg08", i))) == 0) {
        //						ErrNum = 1;
        //						goto PH_PY015_MatrixSpaceLineDel_Error;
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg05", i)) = "" Then '성명
        //						//    ErrNum = 2
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg03", i)) = "" Then '기준년도
        //						//    ErrNum = 3
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg13", i)) = "" Then '구분
        //						//    ErrNum = 4
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg15", i)) = "" Then '교육명
        //						//    ErrNum = 5
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColDt01", i)) = "" Then '교육시작일
        //						//    ErrNum = 6
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColDt02", i)) = "" Then '교육종료일
        //						//    ErrNum = 7
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //						//ElseIf Trim(oDS_PH_PY015B.GetValue("U_ColReg06", i)) = "" Then '부서
        //						//    ErrNum = 8
        //						//    GoTo PH_PY015_MatrixSpaceLineDel_Error
        //					}

        //				}
        //			}

        //			functionReturnValue = true;
        //			return functionReturnValue;
        //			PH_PY015_MatrixSpaceLineDel_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //			if (ErrNum == 1) {
        //				//MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사번이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 2 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 성명이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 3 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 기준년도가 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 4 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 구분이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 5 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육명이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 6 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육시작일이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 7 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 교육종료일이 없습니다. 확인하세요.", "E"
        //				//    ElseIf ErrNum = 8 Then
        //				//        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 부서코드가 없습니다. 확인하세요.", "E"
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PH_PY015_MatrixSpaceLineDel_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}

        //			functionReturnValue = false;
        //			return functionReturnValue;

        //		}

        //		private void PH_PY015_Print_Report01()
        //		{
        //			//******************************************************************************
        //			//Function ID : PH_PY015_Print_Report01()
        //			//해당모듈    : PH_PY015
        //			//기능        : 리포트 출력
        //			//인수        : 없음
        //			//반환값      : 없음
        //			//특이사항    : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string DocNum = null;
        //			short ErrNum = 0;
        //			string WinTitle = null;
        //			string ReportName = null;
        //			string sQry = null;
        //			SAPbobsCOM.Recordset oRecordSet = null;

        //			string CLTCOD = null;
        //			string DestNo1 = null;
        //			string DestNo2 = null;

        //			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			/// ODBC 연결 체크
        //			if (ConnectODBC() == false) {
        //				goto PH_PY015_Print_Report01_Error;
        //			}

        //			////인자 MOVE , Trim 시키기..
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DestNo1 = Strings.Trim(oForm.Items.Item("DestNo1").Specific.VALUE);
        //			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			DestNo2 = Strings.Trim(oForm.Items.Item("DestNo2").Specific.VALUE);

        //			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

        //			WinTitle = "[PH_PY015]";

        //			//창원
        //			if (CLTCOD == "1") {
        //				ReportName = "PH_PY015_01.rpt";
        //			//동래
        //			} else if (CLTCOD == "2") {
        //				ReportName = "PH_PY015_02.rpt";
        //			//사상
        //			} else if (CLTCOD == "3") {
        //				ReportName = "PH_PY015_03.rpt";
        //			}
        //			MDC_Globals.gRpt_Formula = new string[3];
        //			MDC_Globals.gRpt_Formula_Value = new string[3];
        //			MDC_Globals.gRpt_SRptSqry = new string[2];
        //			MDC_Globals.gRpt_SRptName = new string[2];
        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //			//// Formula 수식필드

        //			//// SubReport


        //			MDC_Globals.gRpt_SFormula[1, 1] = "";
        //			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //			/// Procedure 실행"
        //			sQry = "EXEC [PH_PY015_90] '" + CLTCOD + "','" + DestNo1 + "','" + DestNo2 + "'";

        //			//    oRecordSet.DoQuery sQry
        //			//    If oRecordSet.RecordCount = 0 Then
        //			//        ErrNum = 1
        //			//        GoTo PH_PY015_Print_Report01_Error
        //			//    End If

        //			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}

        //			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet = null;
        //			return;
        //			PH_PY015_Print_Report01_Error:

        //			if (ErrNum == 1) {
        //				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oRecordSet = null;
        //				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
        //			} else {
        //				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oRecordSet = null;
        //				MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY015_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //			}

        //		}
        #endregion
    }
}
