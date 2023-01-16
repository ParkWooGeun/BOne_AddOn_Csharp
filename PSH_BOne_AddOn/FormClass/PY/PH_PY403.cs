using System;
using SAPbouiCOM;
using System.Collections.Generic;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정산공제대상자정보등록
    /// </summary>
    internal class PH_PY403 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat;
        private SAPbouiCOM.DBDataSource oDS_PH_PY403A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY403B;
        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        // 전역변수 (정산기초자료 Insert Update 시 사용)
        string CLTCOD;   // 사업장
        string YY;       // 년도
        string CntcCode; // 사번

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY403.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY403_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY403");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";
                
                oForm.Freeze(true);
                PH_PY403_CreateItems();
                PH_PY403_ComboBox_Setting();
                PH_PY403_EnableMenus();
                PH_PY403_FormItemEnabled();
                PH_PY403_AddMatrixRow();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.ActiveItem = "CntcName"; //포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY403_CreateItems()
        {
            try
            {
                oDS_PH_PY403A = oForm.DataSources.DBDataSources.Item("@PH_PY403A");
                oDS_PH_PY403B = oForm.DataSources.DBDataSources.Item("@PH_PY403B");
                oMat = oForm.Items.Item("Mat01").Specific;
                oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat.AutoResizeColumns();

                // 정산년도
                oForm.Items.Item("YY").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// ComboBox_Setting
        /// </summary>
        private void PH_PY403_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                //주택구분
                oForm.Items.Item("House").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("House").Specific.ValidValues.Add("0", "무주택");
                oForm.Items.Item("House").Specific.ValidValues.Add("1", "1주택");
                oForm.Items.Item("House").Specific.ValidValues.Add("2", "2주택이상");
                oForm.Items.Item("House").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("House").DisplayDesc = true;

                //세대구분
                oForm.Items.Item("Saede").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Saede").Specific.ValidValues.Add("Y", "세대주");
                oForm.Items.Item("Saede").Specific.ValidValues.Add("N", "세대원");
                oForm.Items.Item("Saede").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Saede").DisplayDesc = true;

                //부녀자공제
                oForm.Items.Item("Woman").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Woman").Specific.ValidValues.Add("Y", "해당");
                oForm.Items.Item("Woman").Specific.ValidValues.Add("N", "비해당");
                oForm.Items.Item("Woman").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Woman").DisplayDesc = true;

                //한무모공제
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("", "선택");
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("Y", "해당");
                oForm.Items.Item("Sparent").Specific.ValidValues.Add("N", "비해당");
                oForm.Items.Item("Sparent").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Sparent").DisplayDesc = true;

                //재직구분
                oForm.DataSources.UserDataSources.Add("Status_1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT statusID, name FROM [OHST] Order By 1 ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Status_1").Specific, "N");

                //관계 oMat
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P121' AND U_UseYN= 'Y' ";
                oRecordSet.DoQuery(sQry);
                while (!oRecordSet.EoF)
                {
                    oMat.Columns.Item("Relate").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet.MoveNext();
                }
                oMat.Columns.Item("Relate").DisplayDesc = true;

                //소득유무 oMat
                oMat.Columns.Item("Soduk").ValidValues.Add("", "선택");
                oMat.Columns.Item("Soduk").ValidValues.Add("Y", "유(O)");
                oMat.Columns.Item("Soduk").ValidValues.Add("N", "무(X)");
                oMat.Columns.Item("Soduk").DisplayDesc = true;

                //장애인코드 oMat
                oMat.Columns.Item("HDCode").ValidValues.Add("", "해당없음");
                oMat.Columns.Item("HDCode").ValidValues.Add("1", "장애인복지법에 따른 장애인");
                oMat.Columns.Item("HDCode").ValidValues.Add("2", "국가유공자등 예우및지원에 관한 법률에 따른 상이자 및 이와 유사한자로서 근로능력이없는자");
                oMat.Columns.Item("HDCode").ValidValues.Add("3", "그 밖에 항시 치료를 요하는 중증환자");
                oMat.Columns.Item("HDCode").DisplayDesc = true;

                //미숙아.선천성이상아 여부
                oMat.Columns.Item("PreBaby").ValidValues.Add("Y", "해당");
                oMat.Columns.Item("PreBaby").ValidValues.Add("N", "비해당");
                oMat.Columns.Item("PreBaby").DisplayDesc = true;

                //건겅보험산정특례자여부
                oMat.Columns.Item("Tukrae").ValidValues.Add("Y", "해당");
                oMat.Columns.Item("Tukrae").ValidValues.Add("N", "비해당");
                oMat.Columns.Item("Tukrae").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY403_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY403_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YY").Enabled = true;
                    oForm.Items.Item("CntcName").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    PH_PY403_FormClear();            //폼 DocEntry 세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.Items.Item("YY").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);
                    oForm.Items.Item("CntcCode").Specific.Value = "";
                    oForm.Items.Item("CntcName").Specific.Value = "";

                    oForm.EnableMenu("1281", true);  //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YY").Enabled = true;
                    oForm.Items.Item("CntcName").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true);  //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YY").Enabled = false;
                    oForm.Items.Item("CntcName").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY403_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY403A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "사업장은 필수입니다.";
                    throw new Exception();
                }
                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("YY").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("YY").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "년도는 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //사번
                if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                {
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    errMessage = "사번은 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //주택구분
                if (string.IsNullOrEmpty(oForm.Items.Item("House").Specific.Value.ToString().Trim()))
                {
                    errMessage = "주택구분은 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //부녀자공제
                if (string.IsNullOrEmpty(oForm.Items.Item("Woman").Specific.Value.ToString().Trim()))
                {
                    errMessage = "부녀자공제는 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //세대구분
                if (string.IsNullOrEmpty(oForm.Items.Item("Saede").Specific.Value.ToString().Trim()))
                {
                    errMessage = "세대구분은 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //한부모공제
                if (string.IsNullOrEmpty(oForm.Items.Item("Sparent").Specific.Value.ToString().Trim()))
                {
                    errMessage = "한부모공제는 필수입니다. 입력하세요.";
                    throw new Exception();
                }
                //자료check 한부모 부녀자
                if (oForm.Items.Item("Sparent").Specific.Value.ToString().Trim() == "Y" && oForm.Items.Item("Woman").Specific.Value.ToString().Trim() == "Y")
                {
                    errMessage = "부녀자공제와 한부모공제는 중복 Check 할 수 없습니다. 확인하세요.";
                    throw new Exception();
                }
                //라인
                if (oMat.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat.VisualRowCount - 1; i++)
                    {
                        if (oMat.Columns.Item("Chk").Cells.Item(i).Specific.Checked == true)
                        {
                            if (string.IsNullOrEmpty(oMat.Columns.Item("KName").Cells.Item(i).Specific.Value.ToString().Trim()))
                            {
                                oMat.Columns.Item("KName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                errMessage = "셩명은 필수입니다. 입력하세요";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat.Columns.Item("Relate").Cells.Item(i).Specific.Value.ToString().Trim()))
                            {
                                errMessage = "관계는 필수입니다. 입력하세요";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat.Columns.Item("JuminNo").Cells.Item(i).Specific.Value.ToString().Trim()))
                            {
                                oMat.Columns.Item("JuminNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                errMessage = "주민번호는 필수입니다. 입력하세요";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat.Columns.Item("Birthday").Cells.Item(i).Specific.Value.ToString().Trim()))
                            {
                                oMat.Columns.Item("Birthday").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                errMessage = "생년월일은 필수입니다. 입력하세요";
                                throw new Exception();
                            }
                            else if (string.IsNullOrEmpty(oMat.Columns.Item("Soduk").Cells.Item(i).Specific.Value.ToString().Trim()))
                            {
                                errMessage = "소득유무는 필수입니다. 입력하세요";
                                throw new Exception();
                            }
                        }
                    }
                }
                else
                {
                    errMessage = "라인 데이터가 없습니다.";
                    throw new Exception();
                }

                oMat.FlushToDataSource();
                if (oDS_PH_PY403B.Size > 1)
                {
                    oDS_PH_PY403B.RemoveRecord(oDS_PH_PY403B.Size - 1);
                }
                oMat.LoadFromDataSource();

                returnValue = true;
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
           
            return returnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY403_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY403'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY403_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                oMat.FlushToDataSource();
                oRow = oMat.VisualRowCount;

                if (oMat.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY403B.Size <= oMat.VisualRowCount)
                        {
                            oDS_PH_PY403B.InsertRecord(oRow);
                        }
                        oDS_PH_PY403B.Offset = oRow;
                        oDS_PH_PY403B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY403B.SetValue("U_Chk", oRow, "N");
                        oDS_PH_PY403B.SetValue("U_KName", oRow, "");
                        oDS_PH_PY403B.SetValue("U_Relate", oRow, "");
                        oDS_PH_PY403B.SetValue("U_JuminNo", oRow, "");
                        oDS_PH_PY403B.SetValue("U_Birthday", oRow, "");
                        oDS_PH_PY403B.SetValue("U_Soduk", oRow, "");
                        oDS_PH_PY403B.SetValue("U_HDCode", oRow, "");
                        oDS_PH_PY403B.SetValue("U_PreBaby", oRow, "N");
                        oDS_PH_PY403B.SetValue("U_Tukrae", oRow, "N");
                        oMat.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY403B.Offset = oRow - 1;
                        oDS_PH_PY403B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY403B.SetValue("U_Chk", oRow - 1, "N");
                        oDS_PH_PY403B.SetValue("U_KName", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_Relate", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_JuminNo", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_Birthday", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_Soduk", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_HDCode", oRow - 1, "");
                        oDS_PH_PY403B.SetValue("U_PreBaby", oRow - 1, "N");
                        oDS_PH_PY403B.SetValue("U_Tukrae", oRow - 1, "N");
                        oMat.LoadFromDataSource();
                    }
                }
                else if (oMat.VisualRowCount == 0)
                {
                    oDS_PH_PY403B.Offset = oRow;
                    oDS_PH_PY403B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY403B.SetValue("U_Chk", oRow, "N");
                    oDS_PH_PY403B.SetValue("U_KName", oRow, "");
                    oDS_PH_PY403B.SetValue("U_Relate", oRow, "");
                    oDS_PH_PY403B.SetValue("U_JuminNo", oRow, "");
                    oDS_PH_PY403B.SetValue("U_Birthday", oRow, "");
                    oDS_PH_PY403B.SetValue("U_Soduk", oRow, "");
                    oDS_PH_PY403B.SetValue("U_HDCode", oRow, "");
                    oDS_PH_PY403B.SetValue("U_PreBaby", oRow, "N");
                    oDS_PH_PY403B.SetValue("U_Tukrae", oRow, "N");
                    oMat.LoadFromDataSource();
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
        /// 자료조회 및 전년자료가져오기
        /// </summary>
        private void PH_PY403_Find()
        {
            int i;
            string CLTCOD;
            string YY;
            string JYY;
            string CntcCode;
            string Birthday = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                JYY = Convert.ToString(Convert.ToDouble(oForm.Items.Item("YY").Specific.Value.ToString().Trim()) - 1);

                if (codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "1" || codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "2")
                {
                    Birthday = "19" + codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 0, 6);
                }
                else if (codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "3" || codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "4")
                {
                    Birthday = "20" + codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 0, 6);
                }
                else if (codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "5" || codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "6")
                {
                    Birthday = "19" + codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 0, 6);
                }
                else if (codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "7" || codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 6, 1) == "8")
                {
                    Birthday = "20" + codeHelpClass.Mid(oForm.Items.Item("Jumin").Specific.Value.ToString().Trim(), 0, 6);
                }

                sQry = "Select DocEnty = DocEntry,";
                sQry += "      Cnt = Count(*)";
                sQry += " From [@PH_PY403A]";
                sQry += "Where U_CLTCOD = '" + CLTCOD + "'";
                sQry += "  And U_YY = '" + YY + "'";
                sQry += "  And U_CntcCode = '" + CntcCode + "'";
                sQry += " group by DocEntry ";
                oRecordSet.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet.Fields.Item("Cnt").Value.ToString().Trim()) > 0) //자료가 있으면 조회
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY403_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oRecordSet.Fields.Item("DocEnty").Value.ToString().Trim();
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else //자료가 없으면 전년자료 가져오기
                {
                    sQry = "select kname, ";
                    sQry += "      relate, ";
                    sQry += "	   juminno, ";
                    sQry += "	   birthymd, ";
                    sQry += "      hdcode = isnull(max(hdcode),'') ";
                    sQry += "from p_seoybase ";
                    sQry += "where saup = '" + CLTCOD + "'";
                    sQry += "  and yyyy = '" + JYY + "'";
                    sQry += "  and sabun = '" + CntcCode + "'";
                    sQry += "  and div IN('10', '20', '70') ";
                    sQry += "group by kname, ";
                    sQry += "         relate, ";
                    sQry += "         juminno, ";
                    sQry += "         birthymd ";
                    sQry += "order by relate ";
                    oRecordSet.DoQuery(sQry);

                    oDS_PH_PY403B.Clear();
                    oMat.LoadFromDataSource();
                    
                    if (oRecordSet.RecordCount >= 1) //전년자료있는사원
                    {
                        for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                        {
                            oDS_PH_PY403B.InsertRecord(i);
                            oDS_PH_PY403B.Offset = i;
                            oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));

                            if (oRecordSet.Fields.Item("relate").Value.ToString().Trim() == "01") //본인
                            {
                                oDS_PH_PY403B.SetValue("U_Chk", i, "Y");
                                oDS_PH_PY403B.SetValue("U_Soduk", i, "Y");
                            }
                            else
                            {
                                oDS_PH_PY403B.SetValue("U_Chk", i, "N");
                                oDS_PH_PY403B.SetValue("U_Soduk", i, "");
                            }
                            
                            oDS_PH_PY403B.SetValue("U_KName", i, oRecordSet.Fields.Item("kname").Value.ToString().Trim());
                            oDS_PH_PY403B.SetValue("U_Relate", i, oRecordSet.Fields.Item("relate").Value.ToString().Trim());
                            oDS_PH_PY403B.SetValue("U_JuminNo", i, oRecordSet.Fields.Item("juminno").Value.ToString().Trim());
                            oDS_PH_PY403B.SetValue("U_Birthday", i, oRecordSet.Fields.Item("birthymd").Value.ToString().Trim());
                            oDS_PH_PY403B.SetValue("U_HDCode", i, oRecordSet.Fields.Item("hdcode").Value.ToString().Trim());
                            oDS_PH_PY403B.SetValue("U_PreBaby", i, "N");
                            oDS_PH_PY403B.SetValue("U_Tukrae", i, "N");
                            oRecordSet.MoveNext();
                        }
                    }
                    else //전년자료없는사원
                    {
                        // 본인사항등록
                        i = 0;
                        oDS_PH_PY403B.InsertRecord(i);
                        oDS_PH_PY403B.Offset = i;
                        oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PH_PY403B.SetValue("U_Chk", i, "Y");
                        oDS_PH_PY403B.SetValue("U_KName", i, oForm.Items.Item("CntcName").Specific.Value.ToString().Trim());
                        oDS_PH_PY403B.SetValue("U_Relate", i, "01");
                        oDS_PH_PY403B.SetValue("U_JuminNo", i, oForm.Items.Item("Jumin").Specific.Value.ToString().Trim());
                        oDS_PH_PY403B.SetValue("U_Birthday", i, Birthday);
                        oDS_PH_PY403B.SetValue("U_Soduk", i, "Y");
                        oDS_PH_PY403B.SetValue("U_HDCode", i, "");
                        oDS_PH_PY403B.SetValue("U_PreBaby", i, "N");
                        oDS_PH_PY403B.SetValue("U_Tukrae", i, "N");
                    }
                    oMat.LoadFromDataSource();
                    PH_PY403_AddMatrixRow();
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PH_PY403_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string CLTCOD;
            string YY;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YY = oForm.Items.Item("YY").Specific.Value.Trim();

                WinTitle = "[PH_PY403] 연말정산공제대상자신고서출력";
                ReportName = "PH_PY403_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                dataPackFormula.Add(new PSH_DataPackClass("@YY", YY));

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@YY", YY)); 

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY710_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY403_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                            else
                            {
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim();
                                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY403_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY403_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                //정산기초자료 생성
                                sQry = "EXEC [PH_PY403_01] '" + CLTCOD + "', '" + YY + "', '" + CntcCode + "'";
                                oRecordSet.DoQuery(sQry);

                                PH_PY403_FormItemEnabled();
                                PH_PY403_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY403_FormItemEnabled();
                                PH_PY403_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim();
                                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

                                //정산기초자료 생성
                                sQry = "EXEC [PH_PY403_01] '" + CLTCOD + "', '" + YY + "', '" + CntcCode + "'";
                                oRecordSet.DoQuery(sQry);

                                PH_PY403_FormItemEnabled();
                            }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                    }
                    else if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "CntcName" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcName").Specific.Value.ToString().Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string errMessage = string.Empty;
            string CLTCOD;
            string CntcCode;
            string CntcName;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

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
                        switch (pVal.ItemUID)
                        {
                            case "CntcName":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();

                                sQry = "Select Code,";
                                sQry += "      FullName = U_FullName,";
                                sQry += "      TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                           From [@PS_HR200L]";
                                sQry += "                          WHERE Code = '1'";
                                sQry += "                            And U_Code = U_TeamCode),''),";
                                sQry += "      RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                           From [@PS_HR200L]";
                                sQry += "                          WHERE Code = '2'";
                                sQry += "                            And U_Code = U_RspCode),''),";
                                sQry += "      Status = U_Status, ";
                                sQry += "      Jumin = U_govID ";
                                sQry += " From [@PH_PY001A]";
                                sQry += "Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += "  And U_status <> '5'"; // 퇴사자 제외
                                sQry += "  and U_FullName = '" + CntcName + "'";
                                oRecordSet.DoQuery(sQry);

                                oDS_PH_PY403A.SetValue("U_CntcCode", 0, oRecordSet.Fields.Item("Code").Value.ToString().Trim()); //이벤트 안탐
                                oForm.Items.Item("Jumin").Specific.Value = oRecordSet.Fields.Item("Jumin").Value.ToString().Trim();
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("Status_1").Specific.Select(oRecordSet.Fields.Item("Status").Value.ToString().Trim());
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && oRecordSet.RecordCount > 0) //ADD모드이고 사번이있는 사원만
                                {
                                    PH_PY403_Find();
                                }
                                break;

                            case "CntcCode":
                                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

                                sQry = "  Select Code,";
                                sQry += "        FullName = U_FullName,";
                                sQry += "        TeamName = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                            WHERE Code = '1'";
                                sQry += "                              And U_Code = U_TeamCode),''),";
                                sQry += "        RspName  = Isnull((SELECT U_CodeNm";
                                sQry += "                             From [@PS_HR200L]";
                                sQry += "                            WHERE Code = '2'";
                                sQry += "                              And U_Code = U_RspCode),''),";
                                sQry += "        Status = U_Status, ";
                                sQry += "        Jumin = U_govID ";
                                sQry += "   From [@PH_PY001A]";
                                sQry += "  Where U_CLTCOD = '" + CLTCOD + "'";
                                sQry += "    and Code = '" + CntcCode + "'";
                                oRecordSet.DoQuery(sQry);

                                oDS_PH_PY403A.SetValue("U_CntcName", 0, oRecordSet.Fields.Item("FullName").Value.ToString().Trim()); //이벤트 안탐
                                oForm.Items.Item("Jumin").Specific.Value = oRecordSet.Fields.Item("Jumin").Value.ToString().Trim();
                                oForm.Items.Item("TeamName").Specific.Value = oRecordSet.Fields.Item("TeamName").Value.ToString().Trim();
                                oForm.Items.Item("RspName").Specific.Value = oRecordSet.Fields.Item("RspName").Value.ToString().Trim();
                                oForm.Items.Item("Status_1").Specific.Select(oRecordSet.Fields.Item("Status").Value.ToString().Trim());
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && oRecordSet.RecordCount > 0) //ADD모드이고 사번이있는 사원만
                                {
                                    PH_PY403_Find();
                                }
                                break;

                            case "Mat01":
                                if (pVal.ColUID == "KName")
                                {
                                    oMat.FlushToDataSource();
                                    oDS_PH_PY403B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                    oMat.LoadFromDataSource();

                                    if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PH_PY403_AddMatrixRow();
                                    }
                                }
                                else if (pVal.ColUID == "JuminNo") // 주민번호
                                {
                                    if (oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim().Length != 13)
                                    {
                                        oMat.Columns.Item("Birthday").Cells.Item(pVal.Row).Specific.Value = "";
                                        errMessage = "주민등록번호의 자리수가 틀립니다. 확인하세요";
                                        throw new Exception();
                                    }
                                    else // 주민등록번호입력시 생년월일 생성
                                    {
                                        if (codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "1" || codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "2")
                                        {
                                            oMat.Columns.Item("Birthday").Cells.Item(pVal.Row).Specific.Value = "19" + codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 0, 6);
                                        }
                                        else if (codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "3" || codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "4")
                                        {
                                            oMat.Columns.Item("Birthday").Cells.Item(pVal.Row).Specific.Value = "20" + codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 0, 6);
                                        }
                                        else if (codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "5" || codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "6")
                                        {
                                            oMat.Columns.Item("Birthday").Cells.Item(pVal.Row).Specific.Value = "19" + codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 0, 6);
                                        }
                                        else if (codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "7" || codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6, 1) == "8")
                                        {
                                            oMat.Columns.Item("Birthday").Cells.Item(pVal.Row).Specific.Value = "20" + codeHelpClass.Mid(oMat.Columns.Item("JuminNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 0, 6);
                                        }
                                    }
                                }

                                oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat.AutoResizeColumns();
                                break;
                        }
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat.LoadFromDataSource();

                    PH_PY403_FormItemEnabled();
                    PH_PY403_AddMatrixRow();
                    oMat.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY403A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY403B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (oMat.RowCount != oMat.VisualRowCount)
                    {
                        oMat.FlushToDataSource();

                        while (i <= oDS_PH_PY403B.Size - 1)
                        {
                            if (string.IsNullOrEmpty(oDS_PH_PY403B.GetValue("U_LineNum", i)))
                            {
                                oDS_PH_PY403B.RemoveRecord(i);
                                i = 0;
                            }
                            else
                            {
                                i += 1;
                            }
                        }

                        for (i = 0; i <= oDS_PH_PY403B.Size; i++)
                        {
                            oDS_PH_PY403B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        }

                        oMat.LoadFromDataSource();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// DELETE_정산기초자료
        /// </summary>
        private void DELETE_P_SEOYBASE()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YY = oForm.Items.Item("YY").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();

                //정산기초자료 생성(삭제)
                sQry = "DELETE FROM p_seoybase ";
                sQry += "Where saup = '" + CLTCOD + "'";
                sQry += "  And yyyy = '" + YY + "'";
                sQry += "  And sabun = '" + CntcCode + "'";
                sQry += "  And InstType = 'A'";
                oRecordSet.DoQuery(sQry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                            //정산기초자료 생성(삭제)
                            DELETE_P_SEOYBASE();
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
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            break;
                        case "1282": //문서추가
                            PH_PY403_FormItemEnabled();
                            PH_PY403_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY403_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1287": //복제
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}

